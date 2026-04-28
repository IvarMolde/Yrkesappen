'use strict';
const express = require('express');
const https = require('https');
const crypto = require('crypto');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  HeadingLevel, AlignmentType, BorderStyle, WidthType, ShadingType,
  LevelFormat, TabStopType, TabStopPosition, PageNumber, Header, Footer,
} = require('docx');
const pptxgen = require('pptxgenjs');
const fs = require('fs');
const path = require('path');

const app = express();
app.use(express.json({ limit: '10mb' }));
app.use(express.static(path.join(__dirname, 'public')));

// ─── Farger fra DESIGN.md ─────────────────────────────────────────────────────
const C = {
  primary:   '005F73',
  secondary: '0A9396',
  accent:    'E9C46A',
  bgLight:   'F8F9FA',
  bgGray:    'E9ECEF',
  textDark:  '1B1B1B',
  textMid:   '495057',
  white:     'FFFFFF',
};

// ─── Cache for generert innhold (token → data) ────────────────────────────────
// In-memory cache – data lever 30 min
const innholdCache = new Map();
const CACHE_TTL_MS = 30 * 60 * 1000;

function lagreInnhold(data) {
  const token = crypto.randomBytes(16).toString('hex');
  innholdCache.set(token, { data, opprettet: Date.now() });
  // Rydd opp gamle entries
  for (const [t, v] of innholdCache.entries()) {
    if (Date.now() - v.opprettet > CACHE_TTL_MS) innholdCache.delete(t);
  }
  return token;
}

function hentInnhold(token) {
  const entry = innholdCache.get(token);
  if (!entry) return null;
  if (Date.now() - entry.opprettet > CACHE_TTL_MS) {
    innholdCache.delete(token);
    return null;
  }
  return entry.data;
}

// ─── Gemini 2.5 Flash via direkte HTTP ────────────────────────────────────────
function callGemini(prompt) {
  return new Promise((resolve, reject) => {
    const apiKey = process.env.GEMINI_API_KEY;
    if (!apiKey) return reject(new Error('GEMINI_API_KEY mangler i miljøvariabler.'));

    const body = JSON.stringify({
      contents: [{ parts: [{ text: prompt }] }],
      generationConfig: { temperature: 0.7, maxOutputTokens: 16384 },
    });

    const options = {
      hostname: 'generativelanguage.googleapis.com',
      path: `/v1beta/models/gemini-2.5-flash:generateContent?key=${apiKey}`,
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Content-Length': Buffer.byteLength(body),
      },
    };

    const req = https.request(options, (res) => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => {
        try {
          const parsed = JSON.parse(data);
          if (parsed.error) return reject(new Error(parsed.error.message));
          const candidate = parsed.candidates && parsed.candidates[0];
          if (!candidate) return reject(new Error('Tomt Gemini-svar'));
          if (candidate.finishReason && candidate.finishReason !== 'STOP') {
            console.error(`⚠️ Gemini finishReason: ${candidate.finishReason}`);
          }
          const text = candidate.content && candidate.content.parts && candidate.content.parts[0] && candidate.content.parts[0].text;
          if (!text) return reject(new Error('Ingen tekst i Gemini-svar'));
          resolve(text);
        } catch (e) {
          reject(new Error('Kunne ikke tolke svar fra Gemini: ' + data.slice(0, 200)));
        }
      });
    });
    req.setTimeout(45000, () => {
      req.destroy();
      reject(new Error('Gemini-kall timet ut etter 45 sekunder'));
    });
    req.on('error', reject);
    req.write(body);
    req.end();
  });
}

// ─── Hovedprompt ──────────────────────────────────────────────────────────────
function buildPrompt(yrke, niva, sprak, plassering, fokus) {
  const nivaMap = {
    A1: 'svært enkelt språk, maks 40–60 ord per tekst, kun presens, korte SVO-setninger, kun kjente hverdagsord',
    A2: 'enkelt språk, 70–100 ord per tekst, presens og preteritum, enkel inversion, enkle fagord i kontekst',
    B1: 'moderat komplekst, 100–140 ord per tekst, presens/preteritum/perfektum, modalverb, leddsetninger, fagterminologi forklares',
    B2: 'avansert, 150–200 ord per tekst, alle tider, passiv, fagterminologi naturlig, sammensatte setninger',
  };
  const grammatikkMap = {
    A1: 'presens av vanlige verb (er, jobber, har), grunnleggende SVO-ordstilling, tall og daglige fraser',
    A2: 'preteritum av sterke og svake verb, inversion/V2-regelen, enkle bindeord (og, men, fordi, når), enkel adjektivbøyning',
    B1: 'perfektum (har jobbet), modalverb + infinitiv (må/kan/skal/bør), leddsetninger med riktig ordstilling, adjektivbøyning i alle former',
    B2: 'passiv (blir + perfektum partisipp), kondisjonalis (ville + infinitiv), relativsetninger (som, der, hvilket), sammensatte substantiv',
  };

  const hjelpeTekst = sprak && sprak !== 'ingen'
    ? `HJELPESPRÅK: ${sprak}.
KRITISK REGEL: ALL oversettelse i "oversettelse"-feltet MÅ være på ${sprak}. Ikke engelsk, ikke norsk, ikke et annet språk – KUN ${sprak}. Dette gjelder ALLE ord i ordlisten uten unntak.
Plassering: ${plassering === 'tosidig'
      ? `oversettelse i parentes etter norsk term direkte i teksten – oversettelsen MÅ være på ${sprak}`
      : `samle alle oversettelser i ordlisten på slutten – alle oversettelser MÅ være på ${sprak}`}.`
    : 'Ingen hjelpespråk.';

  const fokusInstruksjon = fokus
    ? `\nSPESIELT FOKUS FRA LÆREREN: "${fokus}"\nTa hensyn til dette i alle tre fagtekster, i ordlisten og i oppgavene.\n`
    : '';

  return `Du er en erfaren norsklærer og fagpedagog med dyp kjennskap til CEFR-rammeverket. Lag et komplett arbeidshefte om yrket "${yrke}" på norsknivå ${niva}.

CEFR-NIVÅ ${niva}: ${nivaMap[niva]}
GRAMMATIKKFOKUS FOR ${niva}: ${grammatikkMap[niva]}

${hjelpeTekst}
${fokusInstruksjon}

VIKTIG: Svar KUN med gyldig JSON. Ingen markdown, ingen tekst utenfor JSON.

Heftet skal ha denne strukturen der tekster og oppgaver er FLETTET SAMMEN:
1. Tekst 1 → deretter oppgaver til Tekst 1
2. Tekst 2 → deretter oppgaver til Tekst 2
3. Tekst 3 → deretter oppgaver til Tekst 3
4. Avsluttende oppgaver (vokabular, grammatikk, skriv/muntlig)

{
  "yrke": "${yrke}",
  "niva": "${niva}",
  "intro": "2-3 setninger om yrket tilpasset ${niva}-nivå",
  "seksjoner": [
    { "type": "tekst", "nummer": 1, "tittel": "Tekst 1 – [tema]", "innhold": "..." },
    { "type": "oppgave", "nummer": 1, "tilknyttet_tekst": "Tekst 1", "oppgavetype": "leseforståelse", "tittel": "Leseforståelse – Tekst 1", "instruksjon": "Les Tekst 1 og svar på spørsmålene.", "delopgaver": [ { "bokstav": "a", "tekst": "...", "fasit": "Forventet svar i 1-2 setninger" }, { "bokstav": "b", "tekst": "...", "fasit": "..." }, { "bokstav": "c", "tekst": "...", "fasit": "..." }, { "bokstav": "d", "tekst": "...", "fasit": "..." }, { "bokstav": "e", "tekst": "...", "fasit": "..." } ] },
    { "type": "oppgave", "nummer": 2, "tilknyttet_tekst": "Tekst 1", "oppgavetype": "grammatikk", "tittel": "Grammatikk", "instruksjon": "...", "delopgaver": [ { "bokstav": "a", "tekst": "...", "fasit": "..." }, { "bokstav": "b", "tekst": "...", "fasit": "..." }, { "bokstav": "c", "tekst": "...", "fasit": "..." }, { "bokstav": "d", "tekst": "...", "fasit": "..." }, { "bokstav": "e", "tekst": "...", "fasit": "..." } ] },
    { "type": "tekst", "nummer": 2, "tittel": "Tekst 2 – [tema]", "innhold": "..." },
    { "type": "oppgave", "nummer": 3, "tilknyttet_tekst": "Tekst 2", "oppgavetype": "leseforståelse", "tittel": "...", "instruksjon": "...", "delopgaver": [ { "bokstav": "a", "tekst": "...", "fasit": "..." }, ...5 stk ] },
    { "type": "oppgave", "nummer": 4, "tilknyttet_tekst": "Tekst 2", "oppgavetype": "vokabular", "tittel": "...", "instruksjon": "...", "delopgaver": [ ...5 stk med fasit ] },
    { "type": "tekst", "nummer": 3, "tittel": "Tekst 3 – [tema]", "innhold": "..." },
    { "type": "oppgave", "nummer": 5, "tilknyttet_tekst": "Tekst 3", "oppgavetype": "leseforståelse", "tittel": "...", "instruksjon": "...", "delopgaver": [ ...5 stk med fasit ] },
    { "type": "oppgave", "nummer": 6, "tilknyttet_tekst": "Generell", "oppgavetype": "grammatikk", "tittel": "...", "instruksjon": "...", "delopgaver": [ ...5 stk med fasit ] },
    { "type": "oppgave", "nummer": 7, "tilknyttet_tekst": "Generell", "oppgavetype": "vokabular", "tittel": "...", "instruksjon": "...", "delopgaver": [ ...5 stk med fasit ] },
    { "type": "oppgave", "nummer": 8, "tilknyttet_tekst": "Generell", "oppgavetype": "skriv_muntlig", "tittel": "...", "instruksjon": "...", "delopgaver": [ ...5 stk med fasit (eksempelsvar) ] }
  ],
  "ordliste": [
    { "norsk": "en pasient", "forklaring": "..."${sprak && sprak !== 'ingen' ? `, "oversettelse": "kun ${sprak}"` : ''} }
  ],
  "pptx": {
    "nokkelord": ["8 viktige fagord"],
    "hms": ["4 HMS-punkter"],
    "egenskaper": ["5 personlige egenskaper"],
    "arbeidsoppgaver": ["4 arbeidsoppgaver"],
    "utdanning": ["3 utdanningsveier"]
  }
}

STRENGE KRAV:
- Fagtekstene MÅ være i den lengden angitt for ${niva}
- Tekstene skal bli litt mer krevende fra Tekst 1 til Tekst 3
- Alle oppgaver har nøyaktig 5 delopgaver (a–e)
- Alle delopgaver MÅ ha "fasit"-felt med forventet svar/eksempelsvar
- For leseforståelse: fasit er kort svar (1-2 setninger) basert på teksten
- For skriv_muntlig: fasit er et eksempelsvar (3-5 setninger)
- For grammatikk/vokabular: fasit er konkret riktig svar
- Ordlisten: 12–16 ord
- Legg til 10–12 seksjoner totalt

ORDLISTE-FORMAT (KRITISK):
- SUBSTANTIV: ALLTID med ubestemt artikkel («en», «ei», «et»). Eks: «en pasient», «et sykehus»
- VERB: ALLTID i infinitiv med «å». Eks: «å behandle», «å undersøke»
- ADJEKTIV: grunnform. Eks: «rask», «ansvarlig»
- Yrkestitler: ALLTID liten forbokstav i setninger${sprak && sprak !== 'ingen' ? `\n- Hvert "oversettelse"-felt MÅ inneholde KUN ${sprak}` : ''}`;
}

// ─── Grammatikk-prompt ────────────────────────────────────────────────────────
function buildGrammatikkPrompt(yrke, niva, grammatikkFokus) {
  const nivaBeskriv = {
    A1: 'Svært enkle setninger. Maks 6-8 ord per setning. Kun presens. Hverdagsord.',
    A2: 'Enkle setninger. Maks 10 ord. Presens og preteritum. Kjente faguttrykk.',
    B1: 'Moderat komplekse setninger. Presens, preteritum, perfektum. Fagord forklares.',
    B2: 'Komplekse setninger. Alle tider, passiv, relativsetninger. Fagord brukes naturlig.',
  }[niva] || '';

  const erTilfeldig = grammatikkFokus === 'tilfeldig';
  const fokusInstruksjon = erTilfeldig
    ? `Velg selv et passende grammatisk tema for nivå ${niva}, basert på CEFR. Typiske temaer for ${niva}: ${{
        A1: 'presens av vanlige verb, SVO-ordstilling, personlige pronomen, bestemt/ubestemt form',
        A2: 'preteritum av sterke og svake verb, V2-regelen/inversjon, bindeord, adjektivbøyning',
        B1: 'perfektum, modalverb + infinitiv, leddsetninger, refleksive verb',
        B2: 'passiv konstruksjon, kondisjonalis, relativsetninger, sammensatte ord',
      }[niva]}.`
    : `Grammatisk tema: "${grammatikkFokus}". Tilpass forklaringen og ALLE oppgavene nøyaktig til dette temaet på nivå ${niva}.`;

  return `Du er en erfaren norsklærer som lager grammatikkmateriell for voksne innvandrere på CEFR-nivå ${niva}.
Yrke for eksempler: "${yrke}".

${fokusInstruksjon}

NIVÅTILPASNING ${niva}: ${nivaBeskriv}

LAG EN GRAMMATIKKBLOKK med forklaring og 5 oppgaver.

KRAV TIL FORKLARINGEN:
- Pedagogisk og faglig korrekt norsk bokmål
- Forklar: hva regelen er, når den brukes, og gi 2-3 eksempler fra arbeidslivet
- Maks 120 ord

KRAV TIL OPPGAVENE:
- Alle setninger handler om yrket "${yrke}" eller arbeidslivet
- Stigende vanskelighetsgrad (a → e)
- Yrkestitler med liten forbokstav
- Alle fasit MÅ være korrekt norsk bokmål

Svar KUN med gyldig JSON:

{
  "tema": "Kort tema-navn",
  "forklaring": "Pedagogisk forklaring tilpasset ${niva}.",
  "oppgaver": [
    { "nummer": 1, "type": "fyll_inn", "tittel": "Fyll inn riktig form", "instruksjon": "...", "delopgaver": [ { "bokstav": "a", "tekst": "Setning med (infinitiv).", "fasit": "korrekt form" }, ...5 stk ] },
    { "nummer": 2, "type": "multiple_choice", "tittel": "Velg riktig", "instruksjon": "...", "delopgaver": [ { "bokstav": "a", "tekst": "...", "alternativer": ["a","b","c"], "fasit": "riktig" }, ...5 stk ] },
    { "nummer": 3, "type": "ordstilling", "tittel": "Sett ordene i rekkefølge", "instruksjon": "...", "delopgaver": [ { "bokstav": "a", "tekst": "ord1 / ord2 / ord3", "fasit": "Riktig setning." }, ...5 stk ] },
    { "nummer": 4, "type": "matching", "tittel": "Koble sammen", "instruksjon": "...", "delopgaver": [ { "bokstav": "a", "tekst": "kolonne A", "match": "kolonne B" }, ...5 stk ] },
    { "nummer": 5, "type": "korriger", "tittel": "Korriger feilen", "instruksjon": "...", "delopgaver": [ { "bokstav": "a", "tekst": "Setning med feil.", "fasit": "Riktig setning." }, ...5 stk ] }
  ]
}`;
}

// ─── Normalisér yrkenavn ──────────────────────────────────────────────────────
async function normaliserYrke(yrke) {
  const apiKey = process.env.GEMINI_API_KEY;
  const input = yrke.toLowerCase().trim();
  if (!apiKey) return { yrke: input, korrigert: false };

  const prompt = `Du er en norsk rettskrivingsekspert. Korriger dette yrkenavnet til korrekt bokmål med liten forbokstav.

Input: "${input}"

Svar KUN med JSON: { "yrke": "korrigert", "endret": true/false }

Eksempler:
"Sykepleier" → { "yrke": "sykepleier", "endret": true }
"sykeplier" → { "yrke": "sykepleier", "endret": true }
"kokk" → { "yrke": "kokk", "endret": false }`;

  return new Promise((resolve) => {
    const body = JSON.stringify({
      contents: [{ parts: [{ text: prompt }] }],
      generationConfig: { temperature: 0.1, maxOutputTokens: 150 },
    });
    const options = {
      hostname: 'generativelanguage.googleapis.com',
      path: `/v1beta/models/gemini-2.5-flash:generateContent?key=${apiKey}`,
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'Content-Length': Buffer.byteLength(body) },
    };
    const req = https.request(options, (res) => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => {
        try {
          const parsed = JSON.parse(data);
          let tekst = parsed.candidates[0].content.parts[0].text.trim()
            .replace(/^```(?:json)?\s*/i, '').replace(/\s*```\s*$/i, '');
          const result = JSON.parse(tekst);
          const korrigert = (result.yrke || input).toLowerCase().trim();
          resolve({ yrke: korrigert, korrigert: result.endret === true && korrigert !== input });
        } catch (e) {
          resolve({ yrke: input, korrigert: false });
        }
      });
    });
    req.setTimeout(8000, () => { req.destroy(); resolve({ yrke: input, korrigert: false }); });
    req.on('error', () => resolve({ yrke: input, korrigert: false }));
    req.write(body);
    req.end();
  });
}

// ─── DOCX builder ──────────────────────────────────────────────────────────────
async function buildDocx(data, hjelpesprak, plassering, grammatikkData) {
  const { yrke, niva, intro, seksjoner, ordliste } = data;
  const showHelp = hjelpesprak && hjelpesprak !== 'ingen';
  const ordlisteAtEnd = showHelp && plassering === 'slutt';

  const border1 = { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' };
  const allBorders = { top: border1, bottom: border1, left: border1, right: border1 };
  const noBorder  = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
  const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };

  function hLine() {
    return new Paragraph({
      border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: C.primary, space: 1 } },
      spacing: { after: 120 },
      children: [],
    });
  }

  function sectionHeader(text) {
    return [
      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        spacing: { before: 360, after: 80 },
        children: [new TextRun({ text, bold: true, size: 40, color: C.primary, font: 'Calibri' })],
      }),
      hLine(),
    ];
  }

  function tekstHeader(nr, tittel) {
    return [
      new Paragraph({
        spacing: { before: 360, after: 0 },
        shading: { fill: C.primary, type: ShadingType.CLEAR },
        children: [new TextRun({ text: `  📄 Tekst ${nr}  `, bold: true, size: 24, color: C.white, font: 'Calibri' })],
      }),
      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        spacing: { before: 80, after: 80 },
        children: [new TextRun({ text: tittel, bold: true, size: 32, color: C.secondary, font: 'Calibri' })],
      }),
    ];
  }

  function oppgaveHeader(nr, tittel, instruksjon, tilknyttetTekst, oppgavetype) {
    const typeIkon = { leseforståelse: '📖', grammatikk: '✏️', vokabular: '🔤', skriv_muntlig: '💬' }[oppgavetype] || '📝';
    const visTekst = tilknyttetTekst && tilknyttetTekst !== 'Generell';
    return [
      new Paragraph({
        spacing: { before: 300, after: 60 },
        shading: { fill: C.primary, type: ShadingType.CLEAR },
        keepNext: true, keepLines: true,
        children: [new TextRun({ text: `  Oppgave ${nr}  `, bold: true, size: 26, color: C.white, font: 'Calibri' })],
      }),
      ...(visTekst ? [new Paragraph({
        spacing: { before: 40, after: 40 },
        shading: { fill: C.secondary, type: ShadingType.CLEAR },
        keepNext: true, keepLines: true,
        children: [new TextRun({ text: `  ${typeIkon} ${tilknyttetTekst}  `, size: 20, color: C.white, font: 'Calibri' })],
      })] : []),
      new Paragraph({
        spacing: { before: 60, after: 60 },
        keepNext: true, keepLines: true,
        children: [new TextRun({ text: tittel, bold: true, size: 28, color: C.textDark, font: 'Calibri' })],
      }),
      new Paragraph({
        spacing: { after: 120 },
        keepNext: true, keepLines: true,
        children: [new TextRun({ text: instruksjon, italics: true, size: 24, color: C.textMid, font: 'Calibri' })],
      }),
    ];
  }

  function svarLinje() {
    return new Paragraph({
      spacing: { after: 80 },
      children: [new TextRun({ text: '_'.repeat(58), size: 22, color: 'AAAAAA', font: 'Calibri' })],
    });
  }

  const titleBlock = [
    new Paragraph({
      shading: { fill: C.primary, type: ShadingType.CLEAR },
      spacing: { before: 0, after: 0 },
      children: [new TextRun({ text: `  ${yrke.toUpperCase()}  `, bold: true, size: 56, color: C.white, font: 'Calibri' })],
    }),
    new Paragraph({
      shading: { fill: C.secondary, type: ShadingType.CLEAR },
      spacing: { after: 0 },
      children: [
        new TextRun({ text: `  Arbeidshefte – Norsknivå ${niva}`, size: 28, color: C.white, font: 'Calibri' }),
        new TextRun({ text: '   |   Molde voksenopplæringssenter', size: 24, color: C.bgGray, font: 'Calibri' }),
      ],
    }),
    new Paragraph({ spacing: { after: 200 }, children: [] }),
  ];

  const introBlock = [
    ...sectionHeader('Innledning'),
    new Paragraph({ spacing: { after: 200 }, children: [new TextRun({ text: intro, size: 24, font: 'Calibri' })] }),
  ];

  const seksjonerBlock = [];
  let firstText = true;
  for (const seksjon of seksjoner) {
    if (seksjon.type === 'tekst') {
      if (firstText) { seksjonerBlock.push(...sectionHeader('Fagtekster og oppgaver')); firstText = false; }
      seksjonerBlock.push(...tekstHeader(seksjon.nummer, seksjon.tittel));
      seksjonerBlock.push(new Paragraph({ spacing: { after: 240 }, children: [new TextRun({ text: seksjon.innhold, size: 24, font: 'Calibri' })] }));
    } else if (seksjon.type === 'oppgave') {
      const deloRows = seksjon.delopgaver.map((d, i) => {
        const fill = i % 2 === 0 ? C.white : C.bgGray;
        return new TableRow({
          cantSplit: true,
          children: [
            new TableCell({ borders: noBorders, width: { size: 800, type: WidthType.DXA }, shading: { fill, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 60 }, children: [new Paragraph({ children: [new TextRun({ text: `${d.bokstav})`, bold: true, size: 24, color: C.primary, font: 'Calibri' })] })] }),
            new TableCell({ borders: noBorders, width: { size: 8200, type: WidthType.DXA }, shading: { fill, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 60, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: d.tekst, size: 24, font: 'Calibri' })] }), svarLinje()] }),
          ],
        });
      });
      seksjonerBlock.push(
        ...oppgaveHeader(seksjon.nummer, seksjon.tittel, seksjon.instruksjon, seksjon.tilknyttet_tekst, seksjon.oppgavetype),
        new Table({ width: { size: 9000, type: WidthType.DXA }, columnWidths: [800, 8200], rows: deloRows }),
        new Paragraph({ spacing: { after: 120 }, children: [] }),
      );
    }
  }

  const colCount = showHelp && !ordlisteAtEnd ? 3 : 2;
  const colWidths = colCount === 3 ? [2700, 3500, 2800] : [3300, 5700];

  function makeHeaderCell(text, w) {
    return new TableCell({ borders: allBorders, width: { size: w, type: WidthType.DXA }, shading: { fill: C.primary, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text, bold: true, size: 22, color: C.white, font: 'Calibri' })] })] });
  }

  const headerCells = [makeHeaderCell('Norsk', colWidths[0]), makeHeaderCell('Forklaring', colWidths[1])];
  if (colCount === 3) headerCells.push(makeHeaderCell(hjelpesprak, colWidths[2]));

  const ordRows = ordliste.map((o, i) => {
    const fill = i % 2 === 0 ? C.white : C.bgGray;
    const mc = (text, w, opts = {}) => new TableCell({ borders: allBorders, width: { size: w, type: WidthType.DXA }, shading: { fill, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text, size: 22, font: 'Calibri', ...opts })] })] });
    const cells = [mc(o.norsk, colWidths[0], { bold: true, color: C.secondary }), mc(o.forklaring, colWidths[1])];
    if (colCount === 3) cells.push(mc(o.oversettelse || '', colWidths[2], { italics: true }));
    return new TableRow({ children: cells });
  });

  const ordlisteBlock = [
    ...sectionHeader('Viktige ord og uttrykk'),
    new Table({ width: { size: 9000, type: WidthType.DXA }, columnWidths: colWidths, rows: [new TableRow({ children: headerCells }), ...ordRows] }),
    new Paragraph({ spacing: { after: 200 }, children: [] }),
  ];

  const extraOrdliste = ordlisteAtEnd ? [
    ...sectionHeader(`Ordliste – ${hjelpesprak}`),
    new Table({
      width: { size: 9000, type: WidthType.DXA }, columnWidths: [4500, 4500],
      rows: [
        new TableRow({ children: [makeHeaderCell('Norsk', 4500), makeHeaderCell(hjelpesprak, 4500)] }),
        ...ordliste.map((o, i) => {
          const fill = i % 2 === 0 ? C.white : C.bgGray;
          return new TableRow({ children: [
            new TableCell({ borders: allBorders, width: { size: 4500, type: WidthType.DXA }, shading: { fill, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: o.norsk, size: 22, bold: true, color: C.secondary, font: 'Calibri' })] })] }),
            new TableCell({ borders: allBorders, width: { size: 4500, type: WidthType.DXA }, shading: { fill, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: o.oversettelse || '', size: 22, italics: true, font: 'Calibri' })] })] }),
          ]});
        }),
      ],
    }),
  ] : [];

  // Grammatikkblokk
  const grammatikkBlock = [];
  if (grammatikkData && grammatikkData.oppgaver) {
    const typeIkonGram = { fyll_inn: '✏️', multiple_choice: '☑️', ordstilling: '🔀', matching: '🔗', korriger: '🔍' };
    grammatikkBlock.push(...sectionHeader(`Grammatikk: ${grammatikkData.tema}`));
    grammatikkBlock.push(
      new Paragraph({
        spacing: { before: 0, after: 0 },
        shading: { fill: C.primary, type: ShadingType.CLEAR },
        children: [new TextRun({ text: '  📘 Grammatikkforklaring  ', bold: true, size: 24, color: C.white, font: 'Calibri' })],
      }),
      new Paragraph({
        spacing: { before: 0, after: 240 },
        shading: { fill: 'E6F4F6', type: ShadingType.CLEAR },
        border: { left: { style: BorderStyle.SINGLE, size: 12, color: C.secondary, space: 8 } },
        children: [new TextRun({ text: grammatikkData.forklaring, size: 24, font: 'Calibri', color: C.textDark })],
      })
    );

    grammatikkData.oppgaver.forEach((oppg) => {
      const ikon = typeIkonGram[oppg.type] || '📝';
      grammatikkBlock.push(
        new Paragraph({
          spacing: { before: 280, after: 60 },
          shading: { fill: C.secondary, type: ShadingType.CLEAR },
          keepNext: true, keepLines: true,
          children: [new TextRun({ text: `  ${ikon} Oppgave G${oppg.nummer}: ${oppg.tittel}  `, bold: true, size: 24, color: C.white, font: 'Calibri' })],
        }),
        new Paragraph({
          spacing: { before: 60, after: 100 },
          keepNext: true, keepLines: true,
          children: [new TextRun({ text: oppg.instruksjon, italics: true, size: 22, color: C.textMid, font: 'Calibri' })],
        })
      );

      oppg.delopgaver.forEach((d, idx) => {
        const fill = idx % 2 === 0 ? C.white : C.bgGray;
        if (oppg.type === 'multiple_choice' && d.alternativer) {
          grammatikkBlock.push(
            new Paragraph({
              spacing: { before: 60, after: 20 },
              shading: { fill, type: ShadingType.CLEAR },
              keepNext: true, keepLines: true,
              children: [
                new TextRun({ text: `${d.bokstav})  `, bold: true, size: 24, color: C.primary, font: 'Calibri' }),
                new TextRun({ text: d.tekst, size: 24, font: 'Calibri' }),
              ],
            }),
            new Paragraph({
              spacing: { before: 0, after: 20 },
              shading: { fill, type: ShadingType.CLEAR },
              indent: { left: 360 },
              keepNext: true, keepLines: true,
              children: d.alternativer.map((alt, ai) =>
                new TextRun({ text: `  ${['A', 'B', 'C'][ai]}) ${alt}   `, size: 22, font: 'Calibri', color: C.textMid })
              ),
            }),
            svarLinje()
          );
        } else if (oppg.type === 'matching') {
          grammatikkBlock.push(
            new Table({
              width: { size: 9000, type: WidthType.DXA },
              columnWidths: [400, 4000, 4600],
              rows: [new TableRow({
                cantSplit: true,
                children: [
                  new TableCell({ borders: noBorders, width: { size: 400, type: WidthType.DXA }, shading: { fill, type: ShadingType.CLEAR }, margins: { top: 60, bottom: 60, left: 80, right: 40 }, children: [new Paragraph({ children: [new TextRun({ text: `${d.bokstav})`, bold: true, size: 24, color: C.primary, font: 'Calibri' })] })] }),
                  new TableCell({ borders: { right: { style: BorderStyle.SINGLE, size: 4, color: C.bgGray } }, width: { size: 4000, type: WidthType.DXA }, shading: { fill, type: ShadingType.CLEAR }, margins: { top: 60, bottom: 60, left: 60, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: d.tekst, size: 22, font: 'Calibri' })] })] }),
                  new TableCell({ borders: noBorders, width: { size: 4600, type: WidthType.DXA }, shading: { fill: C.bgGray, type: ShadingType.CLEAR }, margins: { top: 60, bottom: 60, left: 120, right: 60 }, children: [new Paragraph({ children: [new TextRun({ text: '→  ___________________________', size: 22, color: 'AAAAAA', font: 'Calibri' })] })] }),
                ],
              })],
            }),
            new Paragraph({ spacing: { after: 20 }, children: [] })
          );
        } else {
          grammatikkBlock.push(
            new Table({
              width: { size: 9000, type: WidthType.DXA },
              columnWidths: [800, 8200],
              rows: [new TableRow({
                cantSplit: true,
                children: [
                  new TableCell({ borders: noBorders, width: { size: 800, type: WidthType.DXA }, shading: { fill, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 60 }, children: [new Paragraph({ children: [new TextRun({ text: `${d.bokstav})`, bold: true, size: 24, color: C.primary, font: 'Calibri' })] })] }),
                  new TableCell({ borders: noBorders, width: { size: 8200, type: WidthType.DXA }, shading: { fill, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 60, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: d.tekst, size: 24, font: 'Calibri' })] }), svarLinje()] }),
                ],
              })],
            }),
            new Paragraph({ spacing: { after: 20 }, children: [] })
          );
        }
      });

      // Fasit
      grammatikkBlock.push(
        new Paragraph({
          spacing: { before: 120, after: 0 },
          shading: { fill: C.bgGray, type: ShadingType.CLEAR },
          children: [new TextRun({ text: '  🔑 Fasit  ', bold: true, size: 20, color: C.textMid, font: 'Calibri' })],
        })
      );
      oppg.delopgaver.forEach((d) => {
        const fasitTekst = oppg.type === 'matching'
          ? `${d.bokstav}) ${d.tekst}  →  ${d.match}`
          : `${d.bokstav}) ${d.fasit}`;
        grammatikkBlock.push(
          new Paragraph({
            spacing: { before: 20, after: 20 },
            shading: { fill: C.bgGray, type: ShadingType.CLEAR },
            indent: { left: 360 },
            children: [new TextRun({ text: fasitTekst, size: 20, color: C.textMid, font: 'Calibri', italics: true })],
          })
        );
      });
      grammatikkBlock.push(new Paragraph({ spacing: { after: 120 }, children: [] }));
    });
  }

  const doc = new Document({
    numbering: { config: [{ reference: 'bullets', levels: [{ level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] }] },
    styles: {
      default: { document: { run: { font: 'Calibri', size: 24 } } },
      paragraphStyles: [
        { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { size: 40, bold: true, font: 'Calibri', color: C.primary }, paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 0 } },
        { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { size: 32, bold: true, font: 'Calibri', color: C.secondary }, paragraph: { spacing: { before: 180, after: 80 }, outlineLevel: 1 } },
      ],
    },
    sections: [{
      properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1134, right: 1134, bottom: 1134, left: 1134 } } },
      headers: {
        default: new Header({ children: [new Paragraph({
          border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: C.primary, space: 1 } },
          tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
          children: [
            new TextRun({ text: `${yrke} – Nivå ${niva}`, size: 18, color: C.textMid, font: 'Calibri' }),
            new TextRun({ text: '\t', size: 18 }),
            new TextRun({ text: 'Molde voksenopplæringssenter', size: 18, color: C.textMid, font: 'Calibri' }),
          ],
        })] }),
      },
      footers: {
        default: new Footer({ children: [new Paragraph({
          border: { top: { style: BorderStyle.SINGLE, size: 4, color: C.bgGray, space: 1 } },
          tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
          children: [
            new TextRun({ text: '© MBO – Molde voksenopplæringssenter', size: 18, color: C.textMid, font: 'Calibri' }),
            new TextRun({ text: '\tSide ', size: 18, color: C.textMid, font: 'Calibri' }),
            new TextRun({ children: [PageNumber.CURRENT], size: 18, color: C.textMid, font: 'Calibri' }),
          ],
        })] }),
      },
      children: [...titleBlock, ...introBlock, ...seksjonerBlock, ...grammatikkBlock, ...ordlisteBlock, ...extraOrdliste],
    }],
  });

  return Packer.toBuffer(doc);
}

// ─── PPTX builder (uendret fra forrige versjon) ───────────────────────────────
async function buildPptx(data, yrke, niva, hjelpesprak, fokus) {
  const { hms, egenskaper, arbeidsoppgaver, utdanning } = data.pptx;
  const seksjoner = data.seksjoner || [];
  const ordliste = data.ordliste || [];
  const showHelp = hjelpesprak && hjelpesprak !== 'ingen';
  const hasFokus = fokus && fokus.trim().length > 0;
  const tekster = seksjoner.filter(s => s.type === 'tekst').slice(0, 3);

  const pres = new pptxgen();
  pres.layout = 'LAYOUT_16x9';
  pres.title = `${yrke} – Norsknivå ${niva}`;
  pres.author = 'Molde voksenopplæringssenter';

  const makeShadow = () => ({ type: 'outer', blur: 8, offset: 3, angle: 135, color: '000000', opacity: 0.13 });

  function darkSlide() {
    const s = pres.addSlide();
    s.background = { color: C.primary };
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.2, fill: { color: C.accent }, line: { color: C.accent } });
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.425, w: 10, h: 0.2, fill: { color: C.accent }, line: { color: C.accent } });
    return s;
  }

  function lightSlide(titleText, subtitle) {
    const s = pres.addSlide();
    s.background = { color: C.bgLight };
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.15, h: 5.625, fill: { color: C.primary }, line: { color: C.primary } });
    s.addShape(pres.shapes.RECTANGLE, { x: 0.15, y: 0, w: 9.85, h: 0.95, fill: { color: C.white }, line: { color: C.white } });
    s.addText(titleText, { x: 0.3, y: 0.05, w: 9.0, h: 0.6, fontSize: 24, bold: true, color: C.primary, fontFace: 'Calibri', align: 'left', valign: 'middle', margin: 0 });
    if (subtitle) s.addText(subtitle, { x: 0.3, y: 0.62, w: 9.0, h: 0.28, fontSize: 12, color: C.textMid, fontFace: 'Calibri', align: 'left', valign: 'top', margin: 0, italic: true });
    s.addShape(pres.shapes.LINE, { x: 0.3, y: 0.97, w: 9.4, h: 0, line: { color: C.secondary, width: 2 } });
    if (hasFokus) {
      s.addShape(pres.shapes.RECTANGLE, { x: 7.8, y: 0.1, w: 2.0, h: 0.38, fill: { color: C.accent }, line: { color: C.accent } });
      s.addText(`Fokus: ${fokus.slice(0, 28)}`, { x: 7.82, y: 0.1, w: 1.96, h: 0.38, fontSize: 9, bold: true, color: C.textDark, fontFace: 'Calibri', align: 'center', valign: 'middle', margin: 0, wrap: true });
    }
    return s;
  }

  function safeText(s, text, x, y, w, h, opts = {}) {
    s.addText(text, { x, y, w, h, fontSize: opts.fontSize || 14, bold: opts.bold || false, italic: opts.italic || false, color: opts.color || C.textDark, fontFace: 'Calibri', align: opts.align || 'left', valign: opts.valign || 'top', wrap: true, shrinkText: true, margin: opts.margin !== undefined ? opts.margin : 6 });
  }

  // Slide 1 – Tittel
  {
    const s = pres.addSlide();
    s.background = { color: C.primary };
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.2, fill: { color: C.accent }, line: { color: C.accent } });
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.425, w: 10, h: 0.2, fill: { color: C.accent }, line: { color: C.accent } });
    s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.2, w: 9, h: 2.5, fill: { color: '000000', transparency: 45 }, line: { color: '000000', transparency: 45 } });
    s.addText(yrke.toUpperCase(), { x: 0.5, y: 1.3, w: 9, h: 1.5, fontSize: 44, bold: true, color: C.white, fontFace: 'Calibri', align: 'center', valign: 'middle', wrap: true, shrinkText: true, margin: 8 });
    s.addText(`Norsknivå ${niva}`, { x: 0.5, y: 2.85, w: 9, h: 0.5, fontSize: 20, color: C.accent, fontFace: 'Calibri', align: 'center', valign: 'middle', margin: 0 });
    s.addText('Molde voksenopplæringssenter – MBO', { x: 0.5, y: 3.55, w: 9, h: 0.4, fontSize: 13, color: C.bgGray, fontFace: 'Calibri', align: 'center', valign: 'middle', margin: 0 });
    if (hasFokus) s.addText(`Fokus: ${fokus}`, { x: 1.5, y: 4.1, w: 7, h: 0.5, fontSize: 13, italic: true, color: C.accent, fontFace: 'Calibri', align: 'center', valign: 'middle', wrap: true, shrinkText: true, margin: 4 });
  }

  // Slide 2 – Hva er dette yrket?
  {
    const s = lightSlide('Hva er dette yrket?', 'Forberedelse til arbeidsheftet');
    const items = arbeidsoppgaver.map((t, idx) => ({ text: t, options: { bullet: true, breakLine: idx < arbeidsoppgaver.length - 1, fontSize: 15, color: C.textDark, fontFace: 'Calibri', paraSpaceAfter: 8 } }));
    s.addText(items, { x: 0.25, y: 1.1, w: 5.8, h: 4.3, valign: 'top', wrap: true, shrinkText: true, margin: 8 });
    s.addShape(pres.shapes.RECTANGLE, { x: 6.3, y: 1.1, w: 3.45, h: 4.3, fill: { color: C.primary }, line: { color: C.primary }, shadow: makeShadow() });
    s.addText(yrke.charAt(0).toUpperCase(), { x: 6.3, y: 1.1, w: 3.45, h: 3.0, fontSize: 110, bold: true, color: C.white, fontFace: 'Calibri', align: 'center', valign: 'middle', margin: 0 });
    safeText(s, yrke, 6.35, 4.1, 3.35, 0.7, { bold: true, color: C.accent, fontSize: 14, align: 'center', valign: 'middle', margin: 4 });
  }

  // Slide 3 – Viktige ord
  {
    const s = lightSlide('Viktige ord og uttrykk', showHelp ? `Med oversettelse til ${hjelpesprak}` : 'Lær ordene før du leser');
    const ord8 = ordliste.slice(0, 8);
    if (showHelp) {
      s.addShape(pres.shapes.RECTANGLE, { x: 0.2, y: 1.05, w: 4.6, h: 0.38, fill: { color: C.primary }, line: { color: C.primary } });
      safeText(s, 'Norsk', 0.25, 1.05, 4.5, 0.38, { bold: true, color: C.white, align: 'left', valign: 'middle', fontSize: 13, margin: 4 });
      s.addShape(pres.shapes.RECTANGLE, { x: 5.0, y: 1.05, w: 4.6, h: 0.38, fill: { color: C.secondary }, line: { color: C.secondary } });
      safeText(s, hjelpesprak, 5.05, 1.05, 4.5, 0.38, { bold: true, color: C.white, align: 'left', valign: 'middle', fontSize: 13, margin: 4 });
      ord8.forEach((o, i) => {
        const fill = i % 2 === 0 ? C.white : C.bgGray;
        const y = 1.48 + i * 0.5;
        s.addShape(pres.shapes.RECTANGLE, { x: 0.2, y, w: 4.6, h: 0.46, fill: { color: fill }, line: { color: 'DDDDDD', width: 0.5 } });
        safeText(s, o.norsk, 0.28, y, 4.45, 0.46, { bold: true, color: C.secondary, fontSize: 13, valign: 'middle', margin: 4 });
        s.addShape(pres.shapes.RECTANGLE, { x: 5.0, y, w: 4.6, h: 0.46, fill: { color: fill }, line: { color: 'DDDDDD', width: 0.5 } });
        safeText(s, o.oversettelse || o.forklaring || '', 5.08, y, 4.45, 0.46, { italic: true, color: C.textDark, fontSize: 12, valign: 'middle', margin: 4 });
      });
    } else {
      ord8.forEach((o, i) => {
        const col = i % 2; const row = Math.floor(i / 2);
        const x = 0.25 + col * 4.8; const y = 1.1 + row * 1.12;
        s.addShape(pres.shapes.RECTANGLE, { x, y, w: 4.6, h: 1.0, fill: { color: C.white }, line: { color: C.bgGray, width: 1 }, shadow: makeShadow() });
        s.addShape(pres.shapes.RECTANGLE, { x, y, w: 0.12, h: 1.0, fill: { color: C.secondary }, line: { color: C.secondary } });
        safeText(s, o.norsk, x + 0.18, y, 4.35, 0.42, { bold: true, color: C.secondary, fontSize: 14, valign: 'middle', margin: 4 });
        safeText(s, o.forklaring || '', x + 0.18, y + 0.42, 4.35, 0.52, { color: C.textMid, fontSize: 12, valign: 'top', margin: 4 });
      });
    }
  }

  // Slides 4-6
  tekster.forEach((tekst, i) => {
    const s = lightSlide(tekst.tittel, `Tekst ${tekst.nummer} av 3 – forkunnskaper`);
    const aktiviteter = [
      { ikon: '🤔', label: 'Tenk – Par – Del', sporsmal: `Hva tror du en ${yrke} gjør i løpet av en arbeidsdag?`, tips: ['Tenk selv i 30 sekunder', 'Snakk med sidemannen din', 'Del svaret med klassen'] },
      { ikon: '💡', label: 'Brainstorm', sporsmal: `Hvilket utstyr eller hvilke verktøy tror du en ${yrke} bruker?`, tips: ['Skriv ned så mange ord du kan', 'Sammenlign med sidemannen', 'Hvilke ord kjenner du fra før?'] },
      { ikon: '🤝', label: 'Del erfaring', sporsmal: 'Hva er viktig når man jobber sammen med andre?', tips: ['Tenk på din egen arbeidserfaring', 'Hva er god kommunikasjon på jobb?', 'Hva skjer hvis sikkerhetsregler ikke følges?'] },
    ][i] || { ikon: '💬', label: 'Diskuter', sporsmal: `Hva vet du om yrket ${yrke}?`, tips: ['Snakk med sidemannen', 'Del med klassen'] };

    s.addShape(pres.shapes.RECTANGLE, { x: 0.2, y: 1.1, w: 5.9, h: 4.3, fill: { color: C.white }, line: { color: C.bgGray, width: 1 }, shadow: makeShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x: 0.2, y: 1.1, w: 0.14, h: 4.3, fill: { color: C.accent }, line: { color: C.accent } });
    s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 1.15, w: 2.5, h: 0.38, fill: { color: C.accent }, line: { color: C.accent } });
    s.addText(`${aktiviteter.ikon} ${aktiviteter.label}`, { x: 0.4, y: 1.15, w: 2.5, h: 0.38, fontSize: 12, bold: true, color: C.textDark, fontFace: 'Calibri', align: 'center', valign: 'middle', margin: 4 });
    safeText(s, aktiviteter.sporsmal, 0.4, 1.62, 5.7, 1.4, { fontSize: 16, bold: true, color: C.primary, valign: 'middle', align: 'left', margin: 8 });
    s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 3.1, w: 5.7, h: 0.32, fill: { color: C.bgGray }, line: { color: C.bgGray } });
    s.addText('Slik gjør dere det:', { x: 0.45, y: 3.1, w: 5.6, h: 0.32, fontSize: 11, bold: true, color: C.textMid, fontFace: 'Calibri', align: 'left', valign: 'middle', margin: 4 });
    aktiviteter.tips.forEach((tip, j) => {
      const ty = 3.47 + j * 0.6;
      s.addShape(pres.shapes.OVAL, { x: 0.45, y: ty + 0.05, w: 0.38, h: 0.38, fill: { color: C.secondary }, line: { color: C.secondary } });
      s.addText(String(j + 1), { x: 0.45, y: ty + 0.05, w: 0.38, h: 0.38, fontSize: 12, bold: true, color: C.white, fontFace: 'Calibri', align: 'center', valign: 'middle', margin: 0 });
      safeText(s, tip, 0.95, ty, 5.0, 0.5, { fontSize: 13, color: C.textDark, valign: 'middle', margin: 4 });
    });

    s.addShape(pres.shapes.RECTANGLE, { x: 6.3, y: 1.1, w: 3.45, h: 4.3, fill: { color: C.primary }, line: { color: C.primary }, shadow: makeShadow() });
    s.addText(`📖 Tekst ${tekst.nummer}`, { x: 6.35, y: 1.15, w: 3.35, h: 0.45, fontSize: 12, bold: true, color: C.accent, fontFace: 'Calibri', align: 'center', valign: 'middle', margin: 0 });
    s.addText('Nye ord:', { x: 6.35, y: 1.62, w: 3.35, h: 0.32, fontSize: 11, color: C.bgGray, fontFace: 'Calibri', align: 'center', valign: 'middle', italic: true, margin: 0 });
    s.addShape(pres.shapes.LINE, { x: 6.5, y: 2.0, w: 3.1, h: 0, line: { color: C.accent, width: 1 } });
    const relevantOrd = ordliste.slice(i * 3, i * 3 + 4);
    relevantOrd.forEach((o, j) => {
      const oy = 2.1 + j * 0.8;
      s.addShape(pres.shapes.RECTANGLE, { x: 6.4, y: oy, w: 3.25, h: 0.7, fill: { color: C.white }, line: { color: C.bgGray, width: 0.5 } });
      safeText(s, o.norsk, 6.48, oy, 3.1, showHelp ? 0.34 : 0.7, { bold: true, color: C.primary, fontSize: 13, valign: showHelp ? 'top' : 'middle', margin: 5 });
      if (showHelp && o.oversettelse) safeText(s, o.oversettelse, 6.48, oy + 0.34, 3.1, 0.34, { italic: true, color: C.textMid, fontSize: 11, valign: 'top', margin: 4 });
    });
  });

  // Slide 7 – HMS
  {
    const s = lightSlide('Helse, miljø og sikkerhet (HMS)', 'Viktige regler på arbeidsplassen');
    s.addText('HMS', { x: 5.5, y: 0.8, w: 4.2, h: 4.5, fontSize: 130, bold: true, color: C.bgGray, fontFace: 'Calibri', align: 'center', valign: 'middle', margin: 0, transparency: 60 });
    const dotColors = [C.primary, C.secondary, '1A5276', '145A32'];
    hms.slice(0, 4).forEach((punkt, i) => {
      const y = 1.15 + i * 1.08;
      s.addShape(pres.shapes.OVAL, { x: 0.25, y, w: 0.65, h: 0.65, fill: { color: dotColors[i] }, line: { color: dotColors[i] } });
      s.addText(String(i + 1), { x: 0.25, y, w: 0.65, h: 0.65, fontSize: 18, bold: true, color: C.white, fontFace: 'Calibri', align: 'center', valign: 'middle', margin: 0 });
      s.addShape(pres.shapes.RECTANGLE, { x: 1.05, y, w: 5.3, h: 0.9, fill: { color: C.white }, line: { color: C.bgGray, width: 0.5 } });
      safeText(s, punkt, 1.1, y, 5.2, 0.9, { fontSize: 15, color: C.textDark, valign: 'middle', margin: 6 });
    });
  }

  // Slide 8 – Personlige egenskaper
  {
    const s = lightSlide('Personlige egenskaper', 'Hva er viktig for å lykkes?');
    const cardData = [{ col: 0, row: 0 }, { col: 1, row: 0 }, { col: 2, row: 0 }, { col: 0, row: 1 }, { col: 1, row: 1 }];
    egenskaper.slice(0, 5).forEach((eg, i) => {
      const { col, row } = cardData[i];
      const x = 0.25 + col * 3.2; const y = 1.1 + row * 2.1;
      s.addShape(pres.shapes.RECTANGLE, { x, y, w: 3.0, h: 1.9, fill: { color: C.bgGray }, line: { color: C.bgGray }, shadow: makeShadow() });
      s.addShape(pres.shapes.RECTANGLE, { x, y, w: 0.12, h: 1.9, fill: { color: C.secondary }, line: { color: C.secondary } });
      s.addShape(pres.shapes.OVAL, { x: x + 2.5, y: y + 0.08, w: 0.38, h: 0.38, fill: { color: C.primary }, line: { color: C.primary } });
      s.addText(String(i + 1), { x: x + 2.5, y: y + 0.08, w: 0.38, h: 0.38, fontSize: 11, bold: true, color: C.white, fontFace: 'Calibri', align: 'center', valign: 'middle', margin: 0 });
      safeText(s, eg, x + 0.18, y, 2.72, 1.9, { bold: true, color: C.textDark, fontSize: 14, valign: 'middle', align: 'left', margin: 8 });
    });
  }

  // Slide 9 – Utdanning og karriere
  {
    const s = darkSlide();
    s.addText('Utdanning og karriere', { x: 0.4, y: 0.3, w: 9.2, h: 0.7, fontSize: 28, bold: true, color: C.white, fontFace: 'Calibri', align: 'left', valign: 'middle', wrap: true, shrinkText: true, margin: 0 });
    s.addShape(pres.shapes.LINE, { x: 0.4, y: 1.08, w: 9.2, h: 0, line: { color: C.accent, width: 2 } });
    utdanning.slice(0, 3).forEach((u, i) => {
      const x = 0.4 + i * 3.15;
      s.addShape(pres.shapes.RECTANGLE, { x, y: 1.2, w: 3.0, h: 4.0, fill: { color: C.white, transparency: 88 }, line: { color: C.white, width: 1 } });
      s.addText(String(i + 1), { x, y: 1.3, w: 3.0, h: 0.7, fontSize: 34, bold: true, color: C.accent, fontFace: 'Calibri', align: 'center', valign: 'middle', margin: 0 });
      s.addShape(pres.shapes.LINE, { x: x + 0.3, y: 2.05, w: 2.4, h: 0, line: { color: C.accent, width: 1 } });
      safeText(s, u, x + 0.1, 2.15, 2.8, 2.9, { color: C.white, fontSize: 14, valign: 'top', align: 'center', margin: 6 });
    });
  }

  // Slide 10 – La oss snakke norsk!
  {
    const s = lightSlide('La oss snakke norsk! 💬', 'Diskuter med sidepersonen din');
    const sporsmal = [
      `Hva vet du om yrket ${yrke}?`,
      'Ville du likt å jobbe i dette yrket?',
      'Hva er det viktigste å lære for å gjøre denne jobben bra?',
    ];
    sporsmal.forEach((sp, i) => {
      const y = 1.15 + i * 1.45;
      s.addShape(pres.shapes.RECTANGLE, { x: 0.25, y, w: 9.3, h: 1.25, fill: { color: C.white }, line: { color: C.secondary, width: 1.5 }, shadow: makeShadow() });
      s.addShape(pres.shapes.RECTANGLE, { x: 0.25, y, w: 0.5, h: 1.25, fill: { color: C.secondary }, line: { color: C.secondary } });
      s.addText(String(i + 1), { x: 0.25, y, w: 0.5, h: 1.25, fontSize: 22, bold: true, color: C.white, fontFace: 'Calibri', align: 'center', valign: 'middle', margin: 0 });
      safeText(s, sp, 0.85, y, 8.6, 1.25, { fontSize: 17, color: C.primary, bold: true, valign: 'middle', align: 'left', margin: 8 });
    });
  }

  const tmpPath = `/tmp/pptx-${Date.now()}.pptx`;
  await pres.writeFile({ fileName: tmpPath });
  const buf = fs.readFileSync(tmpPath);
  fs.unlinkSync(tmpPath);
  return buf;
}

// ─── HTML builder (interaktiv selvstendig fil) ────────────────────────────────
function escapeHtml(s) {
  return String(s || '')
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

function buildHtml(data, hjelpesprak, plassering, grammatikkData) {
  const { yrke, niva, intro, seksjoner, ordliste } = data;
  const showHelp = hjelpesprak && hjelpesprak !== 'ingen';

  // Embed all data as JSON for client-side
  const klientData = {
    yrke, niva, intro, seksjoner, ordliste,
    grammatikkData: grammatikkData || null,
    hjelpesprak: showHelp ? hjelpesprak : null,
  };

  return `<!DOCTYPE html>
<html lang="no">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>${escapeHtml(yrke)} – Arbeidshefte ${escapeHtml(niva)}</title>
<style>
  *,*::before,*::after { box-sizing: border-box; margin: 0; padding: 0; }
  :root {
    --primary: #005F73; --secondary: #0A9396; --accent: #E9C46A;
    --bgLight: #F8F9FA; --bgGray: #E9ECEF; --textDark: #1B1B1B;
    --textMid: #495057; --white: #FFFFFF;
    --green: #2D9E5C; --red: #C5403A;
  }
  body {
    font-family: 'Segoe UI', Arial, sans-serif;
    background: var(--bgLight); color: var(--textDark);
    line-height: 1.5; padding-bottom: 4rem;
  }
  header {
    background: var(--primary); color: var(--white);
    padding: 2rem 1.5rem; text-align: center;
    border-bottom: 6px solid var(--accent);
  }
  header h1 { font-size: clamp(1.8rem, 4vw, 2.8rem); letter-spacing: -0.5px; }
  header h2 { font-size: 1.1rem; font-weight: 400; opacity: .85; margin-top: .4rem; }
  main { max-width: 900px; margin: 2rem auto; padding: 0 1.5rem; }
  section.kort {
    background: var(--white); border-radius: 12px;
    box-shadow: 0 4px 18px rgba(0,0,0,.08);
    padding: 1.8rem 2rem; margin-bottom: 1.5rem;
  }
  h2.seksjon-tittel {
    color: var(--primary); font-size: 1.6rem; margin-bottom: 1rem;
    border-bottom: 3px solid var(--primary); padding-bottom: .5rem;
  }
  h3.tekst-tittel {
    color: var(--secondary); font-size: 1.3rem; margin-bottom: .8rem;
    display: flex; align-items: center; gap: .5rem;
  }
  .tekst-badge {
    background: var(--primary); color: var(--white);
    padding: .25rem .8rem; border-radius: 6px; font-size: .85rem; font-weight: 700;
  }
  .tekst-innhold { font-size: 1.05rem; line-height: 1.7; color: var(--textDark); }
  .tekst-innhold .ord-link {
    color: var(--secondary); border-bottom: 1px dotted var(--secondary);
    cursor: help; font-weight: 600;
  }
  .oppgave-header {
    background: var(--primary); color: var(--white);
    padding: .6rem 1rem; border-radius: 8px 8px 0 0;
    font-weight: 700; display: flex; align-items: center; gap: .6rem;
  }
  .oppgave-instruksjon {
    font-style: italic; color: var(--textMid);
    margin: .8rem 0; font-size: .95rem;
  }
  .oppg-type-badge {
    background: var(--secondary); color: var(--white);
    padding: .2rem .6rem; border-radius: 4px; font-size: .8rem;
  }
  .delopg {
    margin-bottom: 1rem; padding: .9rem 1rem;
    background: var(--bgLight); border-left: 3px solid var(--bgGray);
    border-radius: 4px; transition: border-color .2s;
  }
  .delopg.korrekt { border-left-color: var(--green); background: #e8f5ed; }
  .delopg.feil { border-left-color: var(--red); background: #fbebeb; }
  .delopg-bokstav {
    display: inline-block; background: var(--primary); color: var(--white);
    width: 28px; height: 28px; line-height: 28px; text-align: center;
    border-radius: 50%; font-weight: 700; margin-right: .6rem;
  }
  .delopg-tekst { display: inline; font-size: 1rem; }
  .svar-input, .svar-textarea {
    width: 100%; margin-top: .6rem; padding: .6rem .8rem;
    border: 2px solid var(--bgGray); border-radius: 6px;
    font-family: inherit; font-size: 1rem; background: var(--white);
    transition: border-color .2s;
  }
  .svar-textarea { min-height: 70px; resize: vertical; }
  .svar-input:focus, .svar-textarea:focus { border-color: var(--secondary); outline: none; }
  .knapp-rad { display: flex; gap: .5rem; margin-top: .6rem; flex-wrap: wrap; }
  .sjekk-btn, .vis-btn {
    background: var(--secondary); color: var(--white);
    border: none; padding: .5rem 1rem; border-radius: 6px;
    font-family: inherit; font-size: .9rem; font-weight: 600;
    cursor: pointer; transition: background .2s;
  }
  .sjekk-btn:hover { background: var(--primary); }
  .vis-btn { background: var(--accent); color: var(--textDark); }
  .vis-btn:hover { background: #d4a93f; }
  .fasit-boks {
    margin-top: .7rem; padding: .7rem .9rem;
    background: #fff8e1; border-left: 4px solid var(--accent);
    border-radius: 4px; font-size: .92rem; line-height: 1.5;
    display: none;
  }
  .fasit-boks.synlig { display: block; }
  .fasit-boks strong { color: var(--primary); }
  .feedback {
    margin-top: .5rem; padding: .5rem .8rem; border-radius: 4px;
    font-size: .9rem; font-weight: 600; display: none;
  }
  .feedback.synlig { display: block; }
  .feedback.ok { background: #d4edda; color: #155724; }
  .feedback.error { background: #f8d7da; color: #721c24; }

  /* Multiple choice */
  .alternativ-rad {
    display: flex; flex-direction: column; gap: .4rem; margin-top: .6rem;
  }
  .alt-knapp {
    background: var(--white); border: 2px solid var(--bgGray);
    padding: .6rem .9rem; border-radius: 6px; cursor: pointer;
    font-family: inherit; font-size: .98rem; text-align: left;
    transition: all .15s; display: flex; align-items: center; gap: .6rem;
  }
  .alt-knapp:hover { border-color: var(--secondary); background: #f0fafa; }
  .alt-knapp.valgt { border-color: var(--primary); background: #e6f4f6; }
  .alt-knapp.korrekt { border-color: var(--green); background: #e8f5ed; }
  .alt-knapp.feil { border-color: var(--red); background: #fbebeb; }
  .alt-knapp .alt-bokstav {
    background: var(--bgGray); width: 26px; height: 26px;
    line-height: 26px; text-align: center; border-radius: 50%;
    font-weight: 700; flex-shrink: 0;
  }

  /* Ordstilling drag-drop */
  .ord-bank, .ord-svar {
    display: flex; flex-wrap: wrap; gap: .5rem; padding: .8rem;
    background: var(--bgLight); border-radius: 6px; min-height: 60px;
    margin-top: .5rem;
  }
  .ord-svar {
    background: #fffbe6; border: 2px dashed var(--accent);
  }
  .ord-svar:empty::before {
    content: 'Klikk på ord under for å sette dem i rekkefølge…';
    color: var(--textMid); font-style: italic; font-size: .88rem;
  }
  .ord-brikke {
    background: var(--white); padding: .4rem .8rem;
    border: 2px solid var(--secondary); border-radius: 6px;
    cursor: pointer; font-size: .95rem; user-select: none;
    transition: all .15s;
  }
  .ord-brikke:hover { background: var(--secondary); color: var(--white); }
  .ord-brikke.brukt { opacity: .35; cursor: default; pointer-events: none; }

  /* Matching */
  .match-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 1rem; margin-top: .8rem; }
  .match-kolonne { display: flex; flex-direction: column; gap: .4rem; }
  .match-item {
    background: var(--white); padding: .6rem .9rem;
    border: 2px solid var(--bgGray); border-radius: 6px;
    cursor: pointer; font-size: .95rem; transition: all .15s;
  }
  .match-item:hover { border-color: var(--secondary); }
  .match-item.valgt { border-color: var(--primary); background: #e6f4f6; }
  .match-item.koblet { border-color: var(--green); background: #e8f5ed; }
  .match-item.koblet::after { content: ' ✓'; color: var(--green); font-weight: 700; }

  /* Ordliste */
  .ordliste-tabell {
    width: 100%; border-collapse: collapse; margin-top: 1rem;
  }
  .ordliste-tabell th {
    background: var(--primary); color: var(--white);
    padding: .7rem; text-align: left; font-size: .9rem;
  }
  .ordliste-tabell td {
    padding: .6rem .8rem; border-bottom: 1px solid var(--bgGray);
    font-size: .95rem;
  }
  .ordliste-tabell tr:nth-child(even) td { background: var(--bgLight); }
  .ordliste-tabell .ord-norsk { color: var(--secondary); font-weight: 600; }

  /* Grammatikk */
  .gram-forklaring {
    background: #e6f4f6; border-left: 4px solid var(--secondary);
    padding: 1rem 1.2rem; border-radius: 0 6px 6px 0; margin: 1rem 0;
  }
  .gram-forklaring h4 { color: var(--primary); margin-bottom: .5rem; }
  .gram-oppgave-header {
    background: var(--secondary);
  }

  /* Fremdriftslinje */
  #fremdrift {
    position: sticky; top: 0; background: var(--white);
    box-shadow: 0 2px 8px rgba(0,0,0,.1); z-index: 10;
    padding: .6rem 1.5rem; display: flex; align-items: center; gap: 1rem;
  }
  #fremdrift-bar {
    flex: 1; height: 8px; background: var(--bgGray);
    border-radius: 4px; overflow: hidden;
  }
  #fremdrift-fyll {
    height: 100%; background: var(--secondary);
    width: 0%; transition: width .3s;
  }
  #fremdrift-tekst { font-size: .85rem; color: var(--textMid); font-weight: 600; }

  footer {
    text-align: center; padding: 2rem; color: var(--textMid);
    font-size: .85rem; border-top: 1px solid var(--bgGray); margin-top: 3rem;
  }
  @media (max-width: 600px) {
    .match-grid { grid-template-columns: 1fr; }
    section.kort { padding: 1.2rem; }
  }
</style>
</head>
<body>

<header>
  <h1>${escapeHtml(yrke.toUpperCase())}</h1>
  <h2>Arbeidshefte – Norsknivå ${escapeHtml(niva)} | Molde voksenopplæringssenter</h2>
</header>

<div id="fremdrift">
  <span id="fremdrift-tekst">Fremgang: 0%</span>
  <div id="fremdrift-bar"><div id="fremdrift-fyll"></div></div>
</div>

<main id="innhold"></main>

<footer>
  © Molde voksenopplæringssenter – MBO &nbsp;|&nbsp; Yrkesappen
</footer>

<script>
const DATA = ${JSON.stringify(klientData)};
let totalOppgaver = 0;
let lostOppgaver = 0;

function oppdaterFremdrift() {
  const pst = totalOppgaver === 0 ? 0 : Math.round((lostOppgaver / totalOppgaver) * 100);
  document.getElementById('fremdrift-fyll').style.width = pst + '%';
  document.getElementById('fremdrift-tekst').textContent = 'Fremgang: ' + pst + '% (' + lostOppgaver + '/' + totalOppgaver + ')';
}

function markerLost(elem) {
  if (elem.dataset.lost === '1') return;
  elem.dataset.lost = '1';
  lostOppgaver++;
  oppdaterFremdrift();
}

function normaliser(s) {
  return String(s || '').toLowerCase().trim().replace(/[.,!?;:]/g, '').replace(/\\s+/g, ' ');
}

function escapeHtml(s) {
  return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;');
}

// ── Bygg fagtekst med klikkbare ordliste-ord ────────────────────────────────
function bygTekstInnhold(innhold) {
  let html = escapeHtml(innhold);
  // Erstatt ordlisteord med klikkbare lenker
  DATA.ordliste.forEach(o => {
    // Strip artikkel/å-merke for matching
    const baseOrd = o.norsk.replace(/^(en |ei |et |å )/i, '');
    if (baseOrd.length < 3) return;
    const regex = new RegExp('\\\\b(' + baseOrd.replace(/[.*+?^\${}()|[\\]\\\\]/g, '\\\\$&') + ')\\\\b', 'gi');
    const tooltip = escapeHtml((o.forklaring || '') + (o.oversettelse ? ' (' + o.oversettelse + ')' : ''));
    html = html.replace(regex, '<span class="ord-link" title="' + tooltip + '">$1</span>');
  });
  return html;
}

// ── Oppgavetype: Leseforståelse / skriv_muntlig (åpent svar) ────────────────
function bygAapenOppgave(seksjon) {
  const div = document.createElement('div');
  div.innerHTML = '<div class="oppgave-header">📝 Oppgave ' + seksjon.nummer + ': ' + escapeHtml(seksjon.tittel) +
    (seksjon.tilknyttet_tekst && seksjon.tilknyttet_tekst !== 'Generell'
      ? '<span class="oppg-type-badge">' + escapeHtml(seksjon.tilknyttet_tekst) + '</span>' : '') + '</div>' +
    '<div class="oppgave-instruksjon">' + escapeHtml(seksjon.instruksjon) + '</div>';

  seksjon.delopgaver.forEach(d => {
    totalOppgaver++;
    const id = 'oppg-' + seksjon.nummer + '-' + d.bokstav;
    const item = document.createElement('div');
    item.className = 'delopg';
    item.innerHTML =
      '<span class="delopg-bokstav">' + d.bokstav + '</span>' +
      '<span class="delopg-tekst">' + escapeHtml(d.tekst) + '</span>' +
      '<textarea class="svar-textarea" id="' + id + '" placeholder="Skriv ditt svar her…"></textarea>' +
      '<div class="knapp-rad">' +
        '<button class="vis-btn" onclick="visEksempel(\\''+id+'\\', this)">📖 Vis eksempelsvar</button>' +
      '</div>' +
      '<div class="fasit-boks" id="fasit-' + id + '"><strong>Eksempelsvar:</strong> ' + escapeHtml(d.fasit || 'Eget svar.') + '</div>';
    item.querySelector('textarea').addEventListener('input', function() {
      if (this.value.trim().length > 5) markerLost(item);
    });
    div.appendChild(item);
  });
  return div;
}

window.visEksempel = function(id, btn) {
  document.getElementById('fasit-' + id).classList.add('synlig');
  btn.disabled = true;
  btn.textContent = '✓ Vist';
  markerLost(btn.closest('.delopg'));
};

// ── Oppgavetype: Multiple choice (grammatikk) ───────────────────────────────
function bygMultipleChoice(d, idx, gramNr) {
  totalOppgaver++;
  const id = 'mc-' + gramNr + '-' + d.bokstav;
  const item = document.createElement('div');
  item.className = 'delopg';
  item.innerHTML =
    '<span class="delopg-bokstav">' + d.bokstav + '</span>' +
    '<span class="delopg-tekst">' + escapeHtml(d.tekst) + '</span>' +
    '<div class="alternativ-rad" id="' + id + '"></div>' +
    '<div class="feedback" id="fb-' + id + '"></div>';
  const altDiv = item.querySelector('.alternativ-rad');
  d.alternativer.forEach((alt, i) => {
    const btn = document.createElement('button');
    btn.className = 'alt-knapp';
    btn.innerHTML = '<span class="alt-bokstav">' + ['A','B','C'][i] + '</span><span>' + escapeHtml(alt) + '</span>';
    btn.onclick = function() {
      // Disable alle
      altDiv.querySelectorAll('.alt-knapp').forEach(b => b.disabled = true);
      const fb = item.querySelector('.feedback');
      if (normaliser(alt) === normaliser(d.fasit)) {
        btn.classList.add('korrekt');
        fb.className = 'feedback ok synlig';
        fb.textContent = '✓ Riktig!';
        item.classList.add('korrekt');
      } else {
        btn.classList.add('feil');
        // Vis riktig
        altDiv.querySelectorAll('.alt-knapp').forEach((b, bi) => {
          if (normaliser(d.alternativer[bi]) === normaliser(d.fasit)) b.classList.add('korrekt');
        });
        fb.className = 'feedback error synlig';
        fb.textContent = '✗ Feil. Riktig svar: ' + d.fasit;
        item.classList.add('feil');
      }
      markerLost(item);
    };
    altDiv.appendChild(btn);
  });
  return item;
}

// ── Oppgavetype: Fyll inn / korriger (sjekk mot fasit) ──────────────────────
function bygFyllInn(d, gramNr, type) {
  totalOppgaver++;
  const id = type + '-' + gramNr + '-' + d.bokstav;
  const item = document.createElement('div');
  item.className = 'delopg';
  item.innerHTML =
    '<span class="delopg-bokstav">' + d.bokstav + '</span>' +
    '<span class="delopg-tekst">' + escapeHtml(d.tekst) + '</span>' +
    '<input type="text" class="svar-input" id="' + id + '" placeholder="Skriv svaret…">' +
    '<div class="knapp-rad">' +
      '<button class="sjekk-btn">Sjekk svar</button>' +
      '<button class="vis-btn">Vis fasit</button>' +
    '</div>' +
    '<div class="feedback" id="fb-' + id + '"></div>' +
    '<div class="fasit-boks" id="fasit-' + id + '"><strong>Fasit:</strong> ' + escapeHtml(d.fasit) + '</div>';
  const inp = item.querySelector('input');
  const fb = item.querySelector('.feedback');
  const btns = item.querySelectorAll('button');
  btns[0].onclick = function() {
    if (normaliser(inp.value) === normaliser(d.fasit)) {
      fb.className = 'feedback ok synlig';
      fb.textContent = '✓ Riktig!';
      item.classList.add('korrekt');
    } else {
      fb.className = 'feedback error synlig';
      fb.textContent = '✗ Ikke helt riktig. Prøv igjen, eller klikk «Vis fasit».';
    }
    markerLost(item);
  };
  btns[1].onclick = function() {
    item.querySelector('.fasit-boks').classList.add('synlig');
    markerLost(item);
  };
  return item;
}

// ── Oppgavetype: Ordstilling (klikk på ord i rekkefølge) ────────────────────
function bygOrdstilling(d, gramNr) {
  totalOppgaver++;
  const id = 'os-' + gramNr + '-' + d.bokstav;
  const item = document.createElement('div');
  item.className = 'delopg';
  const ord = d.tekst.split(/\\s*\\/\\s*/).filter(o => o);
  // Stokk ordene
  const stokket = [...ord].sort(() => Math.random() - 0.5);
  item.innerHTML =
    '<span class="delopg-bokstav">' + d.bokstav + '</span>' +
    '<span class="delopg-tekst">Sett ordene i riktig rekkefølge:</span>' +
    '<div class="ord-svar" id="svar-' + id + '"></div>' +
    '<div class="ord-bank" id="bank-' + id + '"></div>' +
    '<div class="knapp-rad">' +
      '<button class="sjekk-btn">Sjekk</button>' +
      '<button class="vis-btn" onclick="this.previousElementSibling.disabled=true;document.getElementById(\\'fasit-'+id+'\\').classList.add(\\'synlig\\');markerLost(this.closest(\\'.delopg\\'))">Nullstill / vis fasit</button>' +
    '</div>' +
    '<div class="feedback" id="fb-' + id + '"></div>' +
    '<div class="fasit-boks" id="fasit-' + id + '"><strong>Fasit:</strong> ' + escapeHtml(d.fasit) + '</div>';
  const bank = item.querySelector('#bank-' + id);
  const svar = item.querySelector('#svar-' + id);
  stokket.forEach(o => {
    const b = document.createElement('span');
    b.className = 'ord-brikke';
    b.textContent = o;
    b.onclick = function() {
      if (b.classList.contains('brukt')) return;
      const valgt = document.createElement('span');
      valgt.className = 'ord-brikke';
      valgt.textContent = o;
      valgt.onclick = function() {
        b.classList.remove('brukt');
        valgt.remove();
      };
      svar.appendChild(valgt);
      b.classList.add('brukt');
    };
    bank.appendChild(b);
  });
  item.querySelector('.sjekk-btn').onclick = function() {
    const fb = item.querySelector('.feedback');
    const elevenSvar = Array.from(svar.children).map(c => c.textContent).join(' ');
    if (normaliser(elevenSvar) === normaliser(d.fasit)) {
      fb.className = 'feedback ok synlig';
      fb.textContent = '✓ Riktig!';
      item.classList.add('korrekt');
    } else {
      fb.className = 'feedback error synlig';
      fb.textContent = '✗ Ikke riktig. Prøv igjen.';
    }
    markerLost(item);
  };
  return item;
}

// ── Oppgavetype: Matching (klikk-par mellom kolonner) ───────────────────────
function bygMatching(oppg, gramNr) {
  const wrap = document.createElement('div');
  totalOppgaver += oppg.delopgaver.length;
  const venstre = oppg.delopgaver.map(d => ({ tekst: d.tekst, match: d.match, bokstav: d.bokstav }));
  const hoyre = [...oppg.delopgaver].map(d => d.match).sort(() => Math.random() - 0.5);

  wrap.innerHTML = '<div class="match-grid"><div class="match-kolonne" id="m-v-'+gramNr+'"></div><div class="match-kolonne" id="m-h-'+gramNr+'"></div></div><div class="knapp-rad" style="margin-top:1rem;"><button class="vis-btn" onclick="visMatchFasit('+gramNr+')">Vis fasit</button></div><div class="fasit-boks" id="match-fasit-'+gramNr+'"></div>';
  const v = wrap.querySelector('#m-v-' + gramNr);
  const h = wrap.querySelector('#m-h-' + gramNr);
  let valgtVenstre = null;
  const koblet = {};

  venstre.forEach((d, i) => {
    const el = document.createElement('div');
    el.className = 'match-item';
    el.dataset.bokstav = d.bokstav;
    el.dataset.fasit = d.match;
    el.textContent = d.bokstav + ') ' + d.tekst;
    el.onclick = function() {
      if (el.classList.contains('koblet')) return;
      v.querySelectorAll('.match-item').forEach(x => x.classList.remove('valgt'));
      el.classList.add('valgt');
      valgtVenstre = el;
    };
    v.appendChild(el);
  });

  hoyre.forEach((m) => {
    const el = document.createElement('div');
    el.className = 'match-item';
    el.dataset.tekst = m;
    el.textContent = m;
    el.onclick = function() {
      if (el.classList.contains('koblet')) return;
      if (!valgtVenstre) return;
      if (normaliser(valgtVenstre.dataset.fasit) === normaliser(m)) {
        valgtVenstre.classList.remove('valgt');
        valgtVenstre.classList.add('koblet');
        el.classList.add('koblet');
        koblet[valgtVenstre.dataset.bokstav] = m;
        markerLost(valgtVenstre);
        valgtVenstre = null;
      } else {
        el.style.borderColor = 'var(--red)';
        setTimeout(() => el.style.borderColor = '', 600);
      }
    };
    h.appendChild(el);
  });

  // Lagre fasit-data globalt
  window['matchFasit_' + gramNr] = venstre.map(d => d.bokstav + ') ' + d.tekst + ' → ' + d.match).join('<br>');
  return wrap;
}

window.visMatchFasit = function(gramNr) {
  const fb = document.getElementById('match-fasit-' + gramNr);
  fb.innerHTML = '<strong>Fasit:</strong><br>' + window['matchFasit_' + gramNr];
  fb.classList.add('synlig');
};

window.markerLost = markerLost;

// ── Bygg hele dokumentet ────────────────────────────────────────────────────
function byggInnhold() {
  const main = document.getElementById('innhold');

  // Innledning
  const innled = document.createElement('section');
  innled.className = 'kort';
  innled.innerHTML = '<h2 class="seksjon-tittel">Innledning</h2><p>' + escapeHtml(DATA.intro) + '</p>';
  main.appendChild(innled);

  // Tekster og oppgaver
  const tekstSeksjon = document.createElement('section');
  tekstSeksjon.className = 'kort';
  tekstSeksjon.innerHTML = '<h2 class="seksjon-tittel">Fagtekster og oppgaver</h2>';

  DATA.seksjoner.forEach(s => {
    if (s.type === 'tekst') {
      const tekstDiv = document.createElement('div');
      tekstDiv.style.marginTop = '1.5rem';
      tekstDiv.innerHTML =
        '<h3 class="tekst-tittel"><span class="tekst-badge">📄 Tekst ' + s.nummer + '</span> ' + escapeHtml(s.tittel) + '</h3>' +
        '<div class="tekst-innhold">' + bygTekstInnhold(s.innhold) + '</div>';
      tekstSeksjon.appendChild(tekstDiv);
    } else if (s.type === 'oppgave') {
      tekstSeksjon.appendChild(bygAapenOppgave(s));
    }
  });
  main.appendChild(tekstSeksjon);

  // Grammatikkblokk
  if (DATA.grammatikkData) {
    const gd = DATA.grammatikkData;
    const gramSeksjon = document.createElement('section');
    gramSeksjon.className = 'kort';
    gramSeksjon.innerHTML =
      '<h2 class="seksjon-tittel">Grammatikk: ' + escapeHtml(gd.tema) + '</h2>' +
      '<div class="gram-forklaring"><h4>📘 Forklaring</h4><p>' + escapeHtml(gd.forklaring) + '</p></div>';

    gd.oppgaver.forEach(oppg => {
      const ikon = { fyll_inn: '✏️', multiple_choice: '☑️', ordstilling: '🔀', matching: '🔗', korriger: '🔍' }[oppg.type] || '📝';
      const oppgDiv = document.createElement('div');
      oppgDiv.style.marginTop = '1.5rem';
      oppgDiv.innerHTML =
        '<div class="oppgave-header gram-oppgave-header" style="background:var(--secondary)">' +
        ikon + ' Oppgave G' + oppg.nummer + ': ' + escapeHtml(oppg.tittel) + '</div>' +
        '<div class="oppgave-instruksjon">' + escapeHtml(oppg.instruksjon) + '</div>';

      if (oppg.type === 'multiple_choice') {
        oppg.delopgaver.forEach(d => oppgDiv.appendChild(bygMultipleChoice(d, oppg.nummer)));
      } else if (oppg.type === 'fyll_inn' || oppg.type === 'korriger') {
        oppg.delopgaver.forEach(d => oppgDiv.appendChild(bygFyllInn(d, oppg.nummer, oppg.type)));
      } else if (oppg.type === 'ordstilling') {
        oppg.delopgaver.forEach(d => oppgDiv.appendChild(bygOrdstilling(d, oppg.nummer)));
      } else if (oppg.type === 'matching') {
        oppgDiv.appendChild(bygMatching(oppg, oppg.nummer));
      }
      gramSeksjon.appendChild(oppgDiv);
    });
    main.appendChild(gramSeksjon);
  }

  // Ordliste
  const ordSeksjon = document.createElement('section');
  ordSeksjon.className = 'kort';
  let ordHtml = '<h2 class="seksjon-tittel">Viktige ord og uttrykk</h2><table class="ordliste-tabell"><thead><tr><th>Norsk</th><th>Forklaring</th>';
  if (DATA.hjelpesprak) ordHtml += '<th>' + escapeHtml(DATA.hjelpesprak) + '</th>';
  ordHtml += '</tr></thead><tbody>';
  DATA.ordliste.forEach(o => {
    ordHtml += '<tr><td class="ord-norsk">' + escapeHtml(o.norsk) + '</td><td>' + escapeHtml(o.forklaring || '') + '</td>';
    if (DATA.hjelpesprak) ordHtml += '<td><em>' + escapeHtml(o.oversettelse || '') + '</em></td>';
    ordHtml += '</tr>';
  });
  ordHtml += '</tbody></table>';
  ordSeksjon.innerHTML = ordHtml;
  main.appendChild(ordSeksjon);

  oppdaterFremdrift();
}

byggInnhold();
</script>
</body>
</html>`;
}

// ─── API: Logginn ─────────────────────────────────────────────────────────────
app.post('/api/logginn', (req, res) => {
  const { passord } = req.body;
  const riktig = process.env.APP_PASSORD;
  if (!riktig) return res.status(500).json({ ok: false, feil: 'APP_PASSORD ikke satt.' });
  if (passord === riktig) return res.json({ ok: true });
  return res.status(401).json({ ok: false });
});

// ─── API: Generer innhold (returnerer token) ──────────────────────────────────
app.post('/api/generer-tekst', async (req, res) => {
  try {
    const { yrke, niva, sprak, plassering, fokus, grammatikkFokus, passord } = req.body;
    if (!yrke || !niva) return res.status(400).json({ feil: 'Yrke og nivå er påkrevd.' });

    const riktig = process.env.APP_PASSORD;
    if (riktig && passord !== riktig) return res.status(401).json({ feil: 'Ikke autorisert.' });

    const { yrke: yrkeNormalisert, korrigert: yrkeKorrigert } = await normaliserYrke(yrke);
    console.log(`Genererer for: "${yrkeNormalisert}"`);

    const raw = await callGemini(buildPrompt(yrkeNormalisert, niva, sprak, plassering, fokus));
    const clean = raw.trim().replace(/^```(?:json)?\s*/i, '').replace(/\s*```\s*$/i, '');

    let data;
    try {
      data = JSON.parse(clean);
    } catch (e) {
      console.error('JSON feil:', clean.slice(0, 400));
      return res.status(500).json({ feil: 'Klarte ikke tolke svar fra AI. Prøv igjen.' });
    }

    // Rett opp stor forbokstav
    function rettForbokstav(tekst, y) {
      if (!tekst || !y) return tekst;
      const stor = y.charAt(0).toUpperCase() + y.slice(1);
      const liten = y.charAt(0).toLowerCase() + y.slice(1);
      return tekst.replace(
        new RegExp(`(?<![.!?]\\s)(?<!^)\\b${stor}(en|er|ene|ens|s)?\\b`, 'g'),
        (match, ending) => liten + (ending || '')
      );
    }
    if (data.seksjoner) {
      data.seksjoner = data.seksjoner.map(s => {
        if (s.innhold) s.innhold = rettForbokstav(s.innhold, yrkeNormalisert);
        if (s.tittel) s.tittel = rettForbokstav(s.tittel, yrkeNormalisert);
        if (s.instruksjon) s.instruksjon = rettForbokstav(s.instruksjon, yrkeNormalisert);
        if (s.delopgaver) s.delopgaver = s.delopgaver.map(d => ({ ...d, tekst: rettForbokstav(d.tekst, yrkeNormalisert) }));
        return s;
      });
    }
    if (data.intro) data.intro = rettForbokstav(data.intro, yrkeNormalisert);

    // Grammatikkblokk
    let grammatikkData = null;
    const gFokus = (grammatikkFokus || 'ingen').trim();
    if (gFokus !== 'ingen') {
      console.log(`Grammatikk: "${gFokus}"`);
      try {
        const gRaw = await callGemini(buildGrammatikkPrompt(yrkeNormalisert, niva, gFokus));
        const gClean = gRaw.trim().replace(/^```(?:json)?\s*/i, '').replace(/\s*```\s*$/i, '');
        grammatikkData = JSON.parse(gClean);
      } catch (e) {
        console.error('Grammatikk-feil:', e.message);
      }
    }

    // Lagre i cache og returner token
    const token = lagreInnhold({
      data, sprak, plassering, fokus, grammatikkData,
      yrke: yrkeNormalisert, niva,
    });

    res.json({
      ok: true,
      token,
      yrke: yrkeNormalisert,
      niva,
      yrkeKorrigert: yrkeKorrigert ? yrkeNormalisert : null,
      antallTekster: data.seksjoner.filter(s => s.type === 'tekst').length,
      antallOppgaver: data.seksjoner.filter(s => s.type === 'oppgave').length,
      harGrammatikk: !!grammatikkData,
    });

  } catch (err) {
    console.error(err);
    if (!res.headersSent) res.status(500).json({ feil: err.message });
  }
});

// ─── API: Bygg Word ───────────────────────────────────────────────────────────
app.post('/api/build-docx', async (req, res) => {
  try {
    const { token, passord } = req.body;
    const riktig = process.env.APP_PASSORD;
    if (riktig && passord !== riktig) return res.status(401).json({ feil: 'Ikke autorisert.' });

    const c = hentInnhold(token);
    if (!c) return res.status(404).json({ feil: 'Innholdet er utløpt. Generer på nytt.' });

    const buf = await buildDocx(c.data, c.sprak, c.plassering, c.grammatikkData);
    const safeName = c.yrke.replace(/[^a-zA-ZæøåÆØÅ0-9\-]/g, '_');
    res.json({ ok: true, fil: buf.toString('base64'), filnavn: `${safeName}-arbeidshefte-${c.niva}.docx` });
  } catch (err) {
    console.error(err);
    res.status(500).json({ feil: err.message });
  }
});

// ─── API: Bygg PowerPoint ─────────────────────────────────────────────────────
app.post('/api/build-pptx', async (req, res) => {
  try {
    const { token, passord } = req.body;
    const riktig = process.env.APP_PASSORD;
    if (riktig && passord !== riktig) return res.status(401).json({ feil: 'Ikke autorisert.' });

    const c = hentInnhold(token);
    if (!c) return res.status(404).json({ feil: 'Innholdet er utløpt. Generer på nytt.' });

    const buf = await buildPptx(c.data, c.yrke, c.niva, c.sprak, c.fokus);
    const safeName = c.yrke.replace(/[^a-zA-ZæøåÆØÅ0-9\-]/g, '_');
    res.json({ ok: true, fil: buf.toString('base64'), filnavn: `${safeName}-presentasjon-${c.niva}.pptx` });
  } catch (err) {
    console.error(err);
    res.status(500).json({ feil: err.message });
  }
});

// ─── API: Bygg HTML ───────────────────────────────────────────────────────────
app.post('/api/build-html', async (req, res) => {
  try {
    const { token, passord } = req.body;
    const riktig = process.env.APP_PASSORD;
    if (riktig && passord !== riktig) return res.status(401).json({ feil: 'Ikke autorisert.' });

    const c = hentInnhold(token);
    if (!c) return res.status(404).json({ feil: 'Innholdet er utløpt. Generer på nytt.' });

    const html = buildHtml(c.data, c.sprak, c.plassering, c.grammatikkData);
    const safeName = c.yrke.replace(/[^a-zA-ZæøåÆØÅ0-9\-]/g, '_');
    res.json({ ok: true, fil: Buffer.from(html, 'utf8').toString('base64'), filnavn: `${safeName}-interaktiv-${c.niva}.html` });
  } catch (err) {
    console.error(err);
    res.status(500).json({ feil: err.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Yrkesappen kjører på port ${PORT}`));
