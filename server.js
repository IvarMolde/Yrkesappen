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
    { "type": "oppgave", "nummer": 1, "tilknyttet_tekst": "Tekst 1", "oppgavetype": "leseforståelse", "tittel": "Leseforståelse – Tekst 1", "instruksjon": "Les Tekst 1 og løs oppgavene.", "delopgaver": [
        { "bokstav": "a", "oppgtype": "riktig_galt", "tekst": "Påstand basert på teksten.", "fasit": "riktig" },
        { "bokstav": "b", "oppgtype": "fyll_inn", "tekst": "Setning fra teksten med ___ som mangler.", "fasit": "riktig ord" },
        { "bokstav": "c", "oppgtype": "finn_synonym", "tekst": "Finn et ord i teksten som betyr det samme som «[ord]».", "fasit": "synonym fra teksten" },
        { "bokstav": "d", "oppgtype": "multiple_choice", "tekst": "Spørsmål om teksten.", "alternativer": ["alt1","alt2","alt3"], "fasit": "riktig alternativ" },
        { "bokstav": "e", "oppgtype": "skriv_svar", "tekst": "Åpent spørsmål om teksten.", "fasit": "Eksempelsvar.", "alt_fasit": ["Alternativt akseptabelt svar"] }
      ] },
    { "type": "oppgave", "nummer": 2, "tilknyttet_tekst": "Tekst 1", "oppgavetype": "grammatikk", "tittel": "Grammatikk", "instruksjon": "...", "delopgaver": [ { "bokstav": "a", "tekst": "...", "fasit": "..." }, { "bokstav": "b", "tekst": "...", "fasit": "..." }, { "bokstav": "c", "tekst": "...", "fasit": "..." }, { "bokstav": "d", "tekst": "...", "fasit": "..." }, { "bokstav": "e", "tekst": "...", "fasit": "..." } ] },
    { "type": "tekst", "nummer": 2, "tittel": "Tekst 2 – [tema]", "innhold": "..." },
    { "type": "oppgave", "nummer": 3, "tilknyttet_tekst": "Tekst 2", "oppgavetype": "leseforståelse", "tittel": "Leseforståelse – Tekst 2", "instruksjon": "Les Tekst 2 og løs oppgavene.", "delopgaver": [
        { "bokstav": "a", "oppgtype": "riktig_galt", "tekst": "Påstand.", "fasit": "galt" },
        { "bokstav": "b", "oppgtype": "fyll_inn", "tekst": "Setning med ___ .", "fasit": "ord" },
        { "bokstav": "c", "oppgtype": "multiple_choice", "tekst": "Spørsmål.", "alternativer": ["a","b","c"], "fasit": "riktig" },
        { "bokstav": "d", "oppgtype": "finn_synonym", "tekst": "Finn et ord som betyr...", "fasit": "ord" },
        { "bokstav": "e", "oppgtype": "skriv_svar", "tekst": "Åpent spørsmål.", "fasit": "Eksempelsvar." }
      ] },
    { "type": "oppgave", "nummer": 4, "tilknyttet_tekst": "Tekst 2", "oppgavetype": "vokabular", "tittel": "...", "instruksjon": "...", "delopgaver": [ ...5 stk med fasit ] },
    { "type": "tekst", "nummer": 3, "tittel": "Tekst 3 – [tema]", "innhold": "..." },
    { "type": "oppgave", "nummer": 5, "tilknyttet_tekst": "Tekst 3", "oppgavetype": "leseforståelse", "tittel": "Leseforståelse – Tekst 3", "instruksjon": "Les Tekst 3 og løs oppgavene.", "delopgaver": [
        { "bokstav": "a", "oppgtype": "multiple_choice", "tekst": "Spørsmål.", "alternativer": ["a","b","c"], "fasit": "riktig" },
        { "bokstav": "b", "oppgtype": "riktig_galt", "tekst": "Påstand.", "fasit": "riktig" },
        { "bokstav": "c", "oppgtype": "fyll_inn", "tekst": "Setning med ___.", "fasit": "ord" },
        { "bokstav": "d", "oppgtype": "finn_synonym", "tekst": "Finn et ord som betyr...", "fasit": "ord" },
        { "bokstav": "e", "oppgtype": "skriv_svar", "tekst": "Åpent spørsmål.", "fasit": "Eksempelsvar." }
      ] },
    { "type": "oppgave", "nummer": 6, "tilknyttet_tekst": "Generell", "oppgavetype": "grammatikk", "tittel": "...", "instruksjon": "...", "delopgaver": [ ...5 stk med fasit ] },
    { "type": "oppgave", "nummer": 7, "tilknyttet_tekst": "Generell", "oppgavetype": "vokabular", "tittel": "...", "instruksjon": "...", "delopgaver": [ ...5 stk med fasit ] },
    { "type": "oppgave", "nummer": 8, "tilknyttet_tekst": "Generell", "oppgavetype": "vokabular", "tittel": "Ord i kontekst", "instruksjon": "Bruk ordene fra ordlisten i riktig sammenheng.", "delopgaver": [ ...5 stk med fasit ] }
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
- Alle leseforståelsesoppgaver har nøyaktig 5 delopgaver (a–e) med ULIKE oppgavetyper:
  * "oppgtype": "riktig_galt" — påstand som er riktig eller galt. "fasit" MÅ være "riktig" eller "galt"
  * "oppgtype": "fyll_inn" — setning fra teksten med ___ der ett ord mangler. "fasit" er det manglende ordet
  * "oppgtype": "finn_synonym" — finn et ord i teksten som betyr det samme. "fasit" er synonymet fra teksten
  * "oppgtype": "multiple_choice" — spørsmål med 3 alternativer. MÅ ha "alternativer": [...] og "fasit"
  * "oppgtype": "skriv_svar" — åpent spørsmål. "fasit" er et eksempelsvar (1-2 setninger)
- Hver leseforståelsesoppgave MÅ inneholde alle 5 oppgavetyper (én av hver, men i ulik rekkefølge per tekst)
- VIKTIG OM NORSK ORDSTILLING OG INVERSJON:
  * V2-regelen: Verbet skal ALLTID stå på plass 2 i norske hovedsetninger
  * Tidsadverbial og stedsadverbial kan stå først ELLER sist i setningen
  * «I dag jobber hun» og «Hun jobber i dag» er BEGGE grammatisk korrekte
  * «På sykehuset jobber han» og «Han jobber på sykehuset» er BEGGE korrekte
  * FEIL: «I dag hun jobber» (verbet er ikke på plass 2)
  * I leddsetninger er ordstillingen annerledes: «...fordi hun jobber i dag» (subjekt FØR verbal)
  * For ALLE oppgaver der eleven skriver en setning: legg til "alt_fasit": [...] med alternative gyldige formuleringer
  * For "fyll_inn" der svaret er ett enkelt ord: "alt_fasit" er vanligvis ikke nødvendig
- Andre oppgaver (grammatikk, vokabular) har nøyaktig 5 delopgaver (a–e) med "fasit"
- For vokabular: fasit er konkret riktig ord eller uttrykk
- For grammatikk: fasit er konkret riktig svar
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
- ORDSTILLING-OPPGAVER (type "ordstilling"): Husk at norske setninger ofte har flere korrekte ordstillinger:
  * Tidsadverbial kan stå først ELLER sist: «I dag jobber hun» og «Hun jobber i dag» er BEGGE riktige
  * Stedsadverbial kan stå først ELLER sist: «På sykehuset jobber hun» og «Hun jobber på sykehuset»
  * Ved inversjon (V2-regelen) bytter subjekt og verbal plass: «I dag jobber hun» (ikke «I dag hun jobber»)
  * "fasit" = den mest naturlige setningen. "alt_fasit" = array med alle ANDRE gyldige ordstillinger
  * Hvis setningen kun har ÉN korrekt ordstilling, sett "alt_fasit": []

Svar KUN med gyldig JSON:

{
  "tema": "Kort tema-navn",
  "forklaring": "Pedagogisk forklaring tilpasset ${niva}.",
  "oppgaver": [
    { "nummer": 1, "type": "fyll_inn", "tittel": "Fyll inn riktig form", "instruksjon": "...", "delopgaver": [ { "bokstav": "a", "tekst": "Setning med (infinitiv).", "fasit": "korrekt form" }, ...5 stk ] },
    { "nummer": 2, "type": "multiple_choice", "tittel": "Velg riktig", "instruksjon": "...", "delopgaver": [ { "bokstav": "a", "tekst": "...", "alternativer": ["a","b","c"], "fasit": "riktig" }, ...5 stk ] },
    { "nummer": 3, "type": "ordstilling", "tittel": "Sett ordene i rekkefølge", "instruksjon": "...", "delopgaver": [ { "bokstav": "a", "tekst": "ord1 / ord2 / ord3", "fasit": "Hovedsvar.", "alt_fasit": ["Alternativt riktig svar hvis tidsadverbial/stedsadverbial kan flyttes"] }, ...5 stk ] },
    { "nummer": 4, "type": "matching", "tittel": "Koble sammen", "instruksjon": "...", "delopgaver": [ { "bokstav": "a", "tekst": "kolonne A", "match": "kolonne B" }, ...5 stk ] },
    { "nummer": 5, "type": "korriger", "tittel": "Korriger feilen", "instruksjon": "...", "delopgaver": [ { "bokstav": "a", "tekst": "Setning med feil.", "fasit": "Riktig setning.", "alt_fasit": ["Alternativ riktig versjon hvis relevant"] }, ...5 stk ] }
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
    const typeIkon = { leseforståelse: '📖', grammatikk: '✏️', vokabular: '🔤' }[oppgavetype] || '📝';
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
<title>${escapeHtml(yrke)} – Interaktivt arbeidshefte ${escapeHtml(niva)}</title>
<style>
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
:root{--primary:#005F73;--secondary:#0A9396;--accent:#E9C46A;--bgLight:#F8F9FA;--bgGray:#E9ECEF;--textDark:#1B1B1B;--textMid:#495057;--white:#FFFFFF;--green:#2D9E5C;--red:#C5403A}
body{font-family:'Segoe UI',Arial,sans-serif;background:var(--bgLight);color:var(--textDark);min-height:100vh;display:flex;flex-direction:column}
.topbar{background:var(--primary);color:var(--white);display:flex;align-items:center;padding:.6rem 1.2rem;gap:1rem;position:sticky;top:0;z-index:100;box-shadow:0 2px 8px rgba(0,0,0,.15)}
.topbar-tittel{font-weight:700;font-size:1rem;flex-shrink:0;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.topbar-tittel span{color:var(--accent);font-weight:400;margin-left:.3rem}
.score-boks{margin-left:auto;display:flex;align-items:center;gap:.7rem;flex-shrink:0}
.score-ring{width:44px;height:44px;position:relative}
.score-ring svg{transform:rotate(-90deg)}
.score-ring circle{fill:none;stroke-width:4}
.score-ring .bg{stroke:rgba(255,255,255,.2)}
.score-ring .fg{stroke:var(--accent);stroke-linecap:round;transition:stroke-dashoffset .4s}
.score-tall{position:absolute;inset:0;display:flex;align-items:center;justify-content:center;font-size:.8rem;font-weight:800;color:var(--accent)}
.score-label{font-size:.78rem;opacity:.8;line-height:1.2;text-align:right}
.layout{display:flex;flex:1;min-height:0}
.meny{width:240px;background:var(--white);border-right:1px solid var(--bgGray);overflow-y:auto;flex-shrink:0;padding:.8rem 0}
.meny-btn{width:100%;text-align:left;background:none;border:none;padding:.85rem 1.2rem;font-family:inherit;font-size:.92rem;cursor:pointer;display:flex;align-items:center;gap:.7rem;color:var(--textDark);transition:all .15s;border-left:3px solid transparent}
.meny-btn:hover{background:var(--bgLight);color:var(--primary)}
.meny-btn.aktiv{background:#e6f4f6;color:var(--primary);font-weight:700;border-left-color:var(--primary)}
.meny-btn .mi{font-size:1.3rem;flex-shrink:0;width:28px;text-align:center}
.meny-btn .mm{margin-left:auto;background:var(--accent);color:var(--textDark);font-size:.7rem;font-weight:700;padding:.15rem .45rem;border-radius:10px}
.meny-btn .md{margin-left:auto;color:var(--green);font-weight:700;font-size:1rem}
.innhold{flex:1;overflow-y:auto;padding:2rem;max-width:900px}
.side{display:none}
.side.aktiv{display:block}
h2.st{color:var(--primary);font-size:1.5rem;margin-bottom:.5rem;display:flex;align-items:center;gap:.6rem}
h2.st .badge{background:var(--primary);color:var(--white);padding:.2rem .7rem;border-radius:6px;font-size:.8rem}
.si{color:var(--textMid);font-size:.92rem;margin-bottom:1.5rem;font-style:italic}
.tb{background:var(--white);padding:1.5rem;border-radius:10px;box-shadow:0 2px 12px rgba(0,0,0,.06);margin-bottom:1.5rem;line-height:1.7;font-size:1.05rem}
.tb .ol{color:var(--secondary);border-bottom:1px dotted var(--secondary);cursor:help;font-weight:600}
.ok0{background:var(--white);border-radius:10px;box-shadow:0 2px 12px rgba(0,0,0,.06);margin-bottom:1rem;overflow:hidden}
.oh{background:var(--primary);color:var(--white);padding:.7rem 1rem;font-weight:700;font-size:.95rem;display:flex;align-items:center;gap:.5rem}
.oh.g{background:var(--secondary)}
.ob{padding:1.2rem}
.oi{color:var(--textMid);font-style:italic;margin-bottom:1rem;font-size:.9rem}
.d{margin-bottom:1.2rem;padding:.8rem;background:var(--bgLight);border-left:3px solid var(--bgGray);border-radius:4px;transition:border-color .3s,background .3s}
.d.ok{border-left-color:var(--green);background:#e8f5ed}
.d.no{border-left-color:var(--red);background:#fbebeb}
.dl{display:flex;align-items:baseline;gap:.5rem;margin-bottom:.5rem}
.dn{background:var(--primary);color:var(--white);width:26px;height:26px;line-height:26px;text-align:center;border-radius:50%;font-weight:700;font-size:.85rem;flex-shrink:0}
.dt{font-size:1rem}
.si0{width:100%;margin-top:.5rem;padding:.55rem .7rem;border:2px solid var(--bgGray);border-radius:6px;font-family:inherit;font-size:1rem;background:var(--white);transition:border-color .2s}
.si0:focus{border-color:var(--secondary);outline:none}
.br{display:flex;gap:.4rem;margin-top:.5rem;flex-wrap:wrap}
.bs,.bf{border:none;padding:.45rem .9rem;border-radius:6px;font-family:inherit;font-size:.88rem;font-weight:600;cursor:pointer;transition:background .2s}
.bs{background:var(--secondary);color:var(--white)}.bs:hover{background:var(--primary)}
.bf{background:var(--accent);color:var(--textDark)}.bf:hover{background:#d4a93f}
.fb{margin-top:.4rem;padding:.4rem .7rem;border-radius:4px;font-size:.88rem;font-weight:600;display:none}
.fb.v{display:block}.fb.o{background:#d4edda;color:#155724}.fb.e{background:#f8d7da;color:#721c24}
.fx{margin-top:.5rem;padding:.6rem .8rem;background:#fff8e1;border-left:3px solid var(--accent);border-radius:4px;font-size:.9rem;display:none}
.fx.v{display:block}
.al{display:flex;flex-direction:column;gap:.35rem;margin-top:.5rem}
.ab{background:var(--white);border:2px solid var(--bgGray);padding:.55rem .8rem;border-radius:6px;cursor:pointer;font-family:inherit;font-size:.95rem;text-align:left;display:flex;align-items:center;gap:.6rem;transition:all .15s}
.ab:hover:not(:disabled){border-color:var(--secondary);background:#f0fafa}
.ab .al0{background:var(--bgGray);width:24px;height:24px;line-height:24px;text-align:center;border-radius:50%;font-weight:700;font-size:.8rem;flex-shrink:0}
.ab.r{border-color:var(--green);background:#e8f5ed}
.ab.w{border-color:var(--red);background:#fbebeb}
.os{display:flex;flex-wrap:wrap;gap:.4rem;min-height:50px;padding:.6rem;background:#fffbe6;border:2px dashed var(--accent);border-radius:6px;margin-top:.5rem}
.os:empty::before{content:'Klikk ordene nedenfor i riktig rekkefølge\\2026';color:var(--textMid);font-style:italic;font-size:.85rem}
.ob0{display:flex;flex-wrap:wrap;gap:.4rem;margin-top:.5rem}
.op{background:var(--white);padding:.35rem .7rem;border:2px solid var(--secondary);border-radius:6px;cursor:pointer;font-size:.92rem;user-select:none;transition:all .15s}
.op:hover{background:var(--secondary);color:var(--white)}
.op.u{opacity:.3;pointer-events:none}
.ma{display:grid;grid-template-columns:1fr 1fr;gap:1rem;margin-top:.6rem}
.mk{display:flex;flex-direction:column;gap:.35rem}
.mkh{font-size:.8rem;font-weight:700;color:var(--textMid);text-transform:uppercase;margin-bottom:.3rem}
.me{background:var(--white);border:2px solid var(--bgGray);padding:.5rem .7rem;border-radius:6px;cursor:pointer;font-size:.9rem;transition:all .15s}
.me:hover:not(.k){border-color:var(--secondary)}
.me.s{border-color:var(--primary);background:#e6f4f6}
.me.k{border-color:var(--green);background:#e8f5ed;cursor:default}
.me.ff{border-color:var(--red);background:#fbebeb}
.og{display:grid;grid-template-columns:repeat(auto-fill,minmax(260px,1fr));gap:.8rem}
.oc{background:var(--white);border-left:4px solid var(--secondary);border-radius:6px;padding:.8rem 1rem;box-shadow:0 1px 6px rgba(0,0,0,.06)}
.oc .on{color:var(--secondary);font-weight:700;font-size:1rem}
.oc .of{color:var(--textMid);font-size:.88rem;margin-top:.2rem}
.oc .oo{color:var(--primary);font-style:italic;font-size:.85rem;margin-top:.15rem}
.gi{background:#e6f4f6;border-left:4px solid var(--secondary);padding:1rem 1.2rem;border-radius:0 8px 8px 0;margin-bottom:1.5rem}
.gi h4{color:var(--primary);margin-bottom:.4rem}
@media(max-width:700px){
  .layout{flex-direction:column}
  .meny{width:100%;border-right:none;border-bottom:1px solid var(--bgGray);display:flex;overflow-x:auto;padding:.4rem;gap:.3rem}
  .meny-btn{width:auto;white-space:nowrap;padding:.6rem .9rem;border-left:none;border-bottom:3px solid transparent;font-size:.82rem}
  .meny-btn.aktiv{border-left:none;border-bottom-color:var(--primary)}
  .meny-btn .mm,.meny-btn .md{display:none}
  .innhold{padding:1.2rem}
  .ma{grid-template-columns:1fr}
}
</style>
</head>
<body>
<div class="topbar">
  <div class="topbar-tittel">${escapeHtml(yrke.toUpperCase())}<span>Nivå ${escapeHtml(niva)}</span></div>
  <div class="score-boks">
    <div class="score-label"><span id="sT">0 / 0</span><br>poeng</div>
    <div class="score-ring"><svg width="44" height="44" viewBox="0 0 44 44"><circle class="bg" cx="22" cy="22" r="18"/><circle class="fg" id="sR" cx="22" cy="22" r="18" stroke-dasharray="113.1" stroke-dashoffset="113.1"/></svg><div class="score-tall" id="sP">0%</div></div>
  </div>
</div>
<div class="layout">
  <nav class="meny" id="meny"></nav>
  <div class="innhold" id="inn"></div>
</div>
<script>
const D=${JSON.stringify(klientData)};
let p=0,mx=0;
function uS(){const pst=mx===0?0:Math.round(p/mx*100);document.getElementById('sT').textContent=p+' / '+mx;document.getElementById('sP').textContent=pst+'%';const r=document.getElementById('sR');r.style.strokeDashoffset=113.1-(113.1*pst/100);document.querySelectorAll('[data-s]').forEach(b=>{const s=document.getElementById(b.dataset.s);if(!s)return;const t=s.querySelectorAll('.d').length,d=s.querySelectorAll('.d[data-d="1"]').length;const mm=b.querySelector('.mm'),md=b.querySelector('.md');if(t>0&&d===t){if(mm)mm.style.display='none';if(md)md.style.display='inline';}else if(t>0){if(mm){mm.style.display='inline';mm.textContent=d+'/'+t;}if(md)md.style.display='none';}});}
function gP(el,v){if(el.dataset.d==='1')return;el.dataset.d='1';p+=v;uS();}
function n(s){return String(s||'').toLowerCase().trim().replace(/[.,!?;:]/g,'').replace(/\\s+/g,' ');}
function e(s){return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');}
const sd=[];const mn=document.getElementById('meny');const inn=document.getElementById('inn');
function aS(id,ik,tit,fn){const s=document.createElement('div');s.className='side';s.id=id;fn(s);inn.appendChild(s);const b=document.createElement('button');b.className='meny-btn';b.dataset.s=id;b.innerHTML='<span class="mi">'+ik+'</span><span>'+e(tit)+'</span><span class="mm" style="display:none"></span><span class="md" style="display:none">\\u2713</span>';b.onclick=()=>vS(id);mn.appendChild(b);sd.push(id);}
function vS(id){document.querySelectorAll('.side').forEach(s=>s.classList.remove('aktiv'));document.querySelectorAll('.meny-btn').forEach(b=>b.classList.remove('aktiv'));const s=document.getElementById(id);if(s)s.classList.add('aktiv');const b=document.querySelector('[data-s="'+id+'"]');if(b)b.classList.add('aktiv');document.querySelector('.innhold').scrollTop=0;}
function tO(c){let h=e(c);D.ordliste.forEach(o=>{const b=o.norsk.replace(/^(en |ei |et |å )/i,'');if(b.length<3)return;const re=new RegExp('\\\\b('+b.replace(/[.*+?^\${}()|[\\]\\\\]/g,'\\\\$&')+')\\\\b','gi');const t=e((o.forklaring||'')+(o.oversettelse?' ('+o.oversettelse+')':''));h=h.replace(re,'<span class="ol" title="'+t+'">$1</span>');});return h;}

function mkFI(pa,d){mx++;const div=document.createElement('div');div.className='d';
  const alleFasit=[d.fasit];if(d.alt_fasit&&Array.isArray(d.alt_fasit))d.alt_fasit.forEach(a=>{if(a)alleFasit.push(a);});
  const visFasit=alleFasit.join(' / eller: ');
  div.innerHTML='<div class="dl"><span class="dn">'+d.bokstav+'</span><span class="dt">'+e(d.tekst)+'</span></div><input type="text" class="si0" placeholder="Skriv svaret\\u2026"><div class="br"><button class="bs">Sjekk</button><button class="bf">Vis fasit</button></div><div class="fb"></div><div class="fx"><strong>Fasit:</strong> '+e(visFasit)+'</div>';const inp=div.querySelector('input'),fb=div.querySelector('.fb');div.querySelector('.bs').onclick=()=>{const erRiktig=alleFasit.some(f=>n(inp.value)===n(f));if(erRiktig){fb.className='fb o v';fb.textContent='\\u2713 Riktig!';div.classList.add('ok');gP(div,1);}else{fb.className='fb e v';fb.textContent='\\u2717 Ikke riktig. Pr\\u00f8v igjen eller vis fasit.';}};div.querySelector('.bf').onclick=()=>{div.querySelector('.fx').classList.add('v');div.classList.add('no');gP(div,0);};pa.appendChild(div);}

function mkMC(pa,d){mx++;const div=document.createElement('div');div.className='d';div.innerHTML='<div class="dl"><span class="dn">'+d.bokstav+'</span><span class="dt">'+e(d.tekst)+'</span></div><div class="al"></div><div class="fb"></div>';const al=div.querySelector('.al'),fb=div.querySelector('.fb');d.alternativer.forEach((a,i)=>{const b=document.createElement('button');b.className='ab';b.innerHTML='<span class="al0">'+['A','B','C'][i]+'</span><span>'+e(a)+'</span>';b.onclick=()=>{al.querySelectorAll('.ab').forEach(x=>x.disabled=true);if(n(a)===n(d.fasit)){b.classList.add('r');fb.className='fb o v';fb.textContent='\\u2713 Riktig!';div.classList.add('ok');gP(div,1);}else{b.classList.add('w');al.querySelectorAll('.ab').forEach((x,xi)=>{if(n(d.alternativer[xi])===n(d.fasit))x.classList.add('r');});fb.className='fb e v';fb.textContent='\\u2717 Feil. Riktig: '+d.fasit;div.classList.add('no');gP(div,0);}};al.appendChild(b);});pa.appendChild(div);}

function mkRG(pa,d){mx++;const div=document.createElement('div');div.className='d';const riktig=n(d.fasit)==='riktig';div.innerHTML='<div class="dl"><span class="dn">'+d.bokstav+'</span><span class="dt">'+e(d.tekst)+'</span></div><div class="br"><button class="bs" data-v="riktig">\\u2705 Riktig</button><button class="bs" data-v="galt" style="background:var(--accent);color:var(--textDark)">\\u274c Galt</button></div><div class="fb"></div>';const fb=div.querySelector('.fb');div.querySelectorAll('.bs').forEach(b=>{b.onclick=()=>{div.querySelectorAll('.bs').forEach(x=>x.disabled=true);const valgt=b.dataset.v==='riktig';if(valgt===riktig){fb.className='fb o v';fb.textContent='\\u2713 Riktig!';div.classList.add('ok');gP(div,1);}else{fb.className='fb e v';fb.textContent='\\u2717 Feil. P\\u00e5standen er '+(riktig?'riktig':'galt')+'.';div.classList.add('no');gP(div,0);}};});pa.appendChild(div);}

function mkSyn(pa,d){mx++;const div=document.createElement('div');div.className='d';
  const alleFasit=[d.fasit];if(d.alt_fasit&&Array.isArray(d.alt_fasit))d.alt_fasit.forEach(a=>{if(a)alleFasit.push(a);});
  const visFasit=alleFasit.join(' / ');
  div.innerHTML='<div class="dl"><span class="dn">'+d.bokstav+'</span><span class="dt">'+e(d.tekst)+'</span></div><input type="text" class="si0" placeholder="Skriv ordet\\u2026"><div class="br"><button class="bs">Sjekk</button><button class="bf">Vis fasit</button></div><div class="fb"></div><div class="fx"><strong>Fasit:</strong> '+e(visFasit)+'</div>';const inp=div.querySelector('input'),fb=div.querySelector('.fb');div.querySelector('.bs').onclick=()=>{const erRiktig=alleFasit.some(f=>n(inp.value)===n(f));if(erRiktig){fb.className='fb o v';fb.textContent='\\u2713 Riktig!';div.classList.add('ok');gP(div,1);}else{fb.className='fb e v';fb.textContent='\\u2717 Ikke riktig. Pr\\u00f8v igjen.';}};div.querySelector('.bf').onclick=()=>{div.querySelector('.fx').classList.add('v');div.classList.add('no');gP(div,0);};pa.appendChild(div);}

function mkAA(pa,d){mx++;const div=document.createElement('div');div.className='d';div.innerHTML='<div class="dl"><span class="dn">'+d.bokstav+'</span><span class="dt">'+e(d.tekst)+'</span></div><input type="text" class="si0" placeholder="Skriv svaret ditt\\u2026"><div class="br"><button class="bs">Sjekk</button><button class="bf">Vis eksempelsvar</button></div><div class="fb"></div><div class="fx"><strong>Eksempelsvar:</strong> '+e(d.fasit||'')+'</div>';const inp=div.querySelector('input'),fb=div.querySelector('.fb');div.querySelector('.bs').onclick=()=>{const s=n(inp.value);const fw=n(d.fasit).split(' ').filter(w=>w.length>3);const tr=fw.filter(w=>s.includes(w)).length;if(fw.length>0&&tr/fw.length>=0.4){fb.className='fb o v';fb.textContent='\\u2713 Godt svar!';div.classList.add('ok');gP(div,1);}else if(s.length>3){fb.className='fb e v';fb.textContent='Sjekk eksempelsvaret.';div.querySelector('.fx').classList.add('v');gP(div,0);}else{fb.className='fb e v';fb.textContent='Skriv litt mer\\u2026';}};div.querySelector('.bf').onclick=()=>{div.querySelector('.fx').classList.add('v');if(div.dataset.d!=='1')gP(div,0);};pa.appendChild(div);}

function mkOS(pa,d){mx++;const div=document.createElement('div');div.className='d';const ord=d.tekst.split(/\\s*\\/\\s*/).filter(Boolean);const st=[...ord].sort(()=>Math.random()-.5);
  // Samle alle gyldige svar: fasit + alt_fasit
  const alleFasit=[d.fasit];if(d.alt_fasit&&Array.isArray(d.alt_fasit))d.alt_fasit.forEach(a=>{if(a)alleFasit.push(a);});
  const visFasit=alleFasit.join(' / eller: ');
  div.innerHTML='<div class="dl"><span class="dn">'+d.bokstav+'</span><span class="dt">Sett ordene i riktig rekkef\\u00f8lge:</span></div><div class="os"></div><div class="ob0"></div><div class="br"><button class="bs">Sjekk</button><button class="bf">Vis fasit</button></div><div class="fb"></div><div class="fx"><strong>Fasit:</strong> '+e(visFasit)+'</div>';const sv=div.querySelector('.os'),bk=div.querySelector('.ob0'),fb=div.querySelector('.fb');st.forEach(o=>{const b=document.createElement('span');b.className='op';b.textContent=o;b.onclick=()=>{if(b.classList.contains('u'))return;const v=document.createElement('span');v.className='op';v.textContent=o;v.onclick=()=>{b.classList.remove('u');v.remove();};sv.appendChild(v);b.classList.add('u');};bk.appendChild(b);});div.querySelector('.bs').onclick=()=>{const r=Array.from(sv.children).map(c=>c.textContent).join(' ');const erRiktig=alleFasit.some(f=>n(r)===n(f));if(erRiktig){fb.className='fb o v';fb.textContent='\\u2713 Riktig!';div.classList.add('ok');gP(div,1);}else{fb.className='fb e v';fb.textContent='\\u2717 Ikke riktig. Pr\\u00f8v igjen.';}};div.querySelector('.bf').onclick=()=>{div.querySelector('.fx').classList.add('v');div.classList.add('no');gP(div,0);};pa.appendChild(div);}

function mkMA(pa,oppg){const w=document.createElement('div');const hr=[...oppg.delopgaver].map(d=>d.match).sort(()=>Math.random()-.5);w.innerHTML='<div class="ma"><div class="mk"></div><div class="mk"></div></div><div class="br" style="margin-top:.8rem"><button class="bf">Vis fasit</button></div><div class="fx"></div>';const cols=w.querySelectorAll('.mk');const kv=cols[0];const kh=cols[1];kv.innerHTML='<div class="mkh">Kolonne A</div>';kh.innerHTML='<div class="mkh">Kolonne B</div>';let vv=null;oppg.delopgaver.forEach(d=>{mx++;const el=document.createElement('div');el.className='me d';el.dataset.f=d.match;el.textContent=d.bokstav+') '+d.tekst;el.onclick=()=>{if(el.classList.contains('k'))return;kv.querySelectorAll('.me').forEach(x=>x.classList.remove('s'));el.classList.add('s');vv=el;};kv.appendChild(el);});hr.forEach(m=>{const el=document.createElement('div');el.className='me';el.textContent=m;el.onclick=()=>{if(el.classList.contains('k')||!vv)return;if(n(vv.dataset.f)===n(m)){vv.classList.remove('s');vv.classList.add('k','ok');el.classList.add('k');gP(vv,1);vv=null;}else{el.classList.add('ff');setTimeout(()=>el.classList.remove('ff'),500);}};kh.appendChild(el);});const fx=w.querySelector('.fx');fx.innerHTML=oppg.delopgaver.map(d=>'<strong>'+d.bokstav+')</strong> '+e(d.tekst)+' \\u2192 '+e(d.match)).join('<br>');w.querySelector('.bf').onclick=()=>{fx.classList.add('v');kv.querySelectorAll('.me:not(.k)').forEach(el=>{el.classList.add('no');gP(el,0);});};pa.appendChild(w);}

// Smart dispatcher for leseforståelse delopgaver
function mkLese(pa,d){
  const t=d.oppgtype||'skriv_svar';
  if(t==='riktig_galt')mkRG(pa,d);
  else if(t==='multiple_choice'&&d.alternativer)mkMC(pa,d);
  else if(t==='fyll_inn')mkFI(pa,d);
  else if(t==='finn_synonym')mkSyn(pa,d);
  else mkAA(pa,d);
}

// \\u2550 BYGG SIDER \\u2550
const tekster=D.seksjoner.filter(s=>s.type==='tekst');
const leseO=D.seksjoner.filter(s=>s.type==='oppgave'&&s.oppgavetype==='leseforst\\u00e5else');
aS('s-t','\\ud83d\\udcd6','Fagtekst og leseforst\\u00e5else',s=>{
  s.innerHTML='<h2 class="st"><span class="badge">\\ud83d\\udcd6</span> Fagtekst og leseforst\\u00e5else</h2><p class="si">'+e(D.intro)+'</p>';
  tekster.forEach(t=>{
    s.insertAdjacentHTML('beforeend','<h3 style="color:var(--secondary);margin:1.5rem 0 .5rem;font-size:1.15rem">\\ud83d\\udcc4 Tekst '+t.nummer+': '+e(t.tittel)+'</h3><div class="tb">'+tO(t.innhold)+'</div>');
    const tilknyttet=leseO.filter(o=>o.tilknyttet_tekst==='Tekst '+t.nummer);
    tilknyttet.forEach(oppg=>{
      const k=document.createElement('div');
      k.className='ok0';
      k.innerHTML='<div class="oh">Oppgave '+oppg.nummer+': '+e(oppg.tittel)+'</div><div class="ob"><div class="oi">'+e(oppg.instruksjon)+'</div></div>';
      const b=k.querySelector('.ob');
      oppg.delopgaver.forEach(d=>mkLese(b,d));
      s.appendChild(k);
    });
  });
});

const gvO=D.seksjoner.filter(s=>s.type==='oppgave'&&(s.oppgavetype==='grammatikk'||s.oppgavetype==='vokabular'));
if(gvO.length>0){aS('s-gv','\\u270f\\ufe0f','Grammatikk & Vokabular',s=>{
  s.innerHTML='<h2 class="st"><span class="badge">\\u270f\\ufe0f</span> Grammatikk & Vokabular</h2><p class="si">\\u00d8v p\\u00e5 ord og grammatikk fra tekstene.</p>';
  gvO.forEach(oppg=>{
    const k=document.createElement('div');k.className='ok0';
    k.innerHTML='<div class="oh">Oppgave '+oppg.nummer+': '+e(oppg.tittel)+'</div><div class="ob"><div class="oi">'+e(oppg.instruksjon)+'</div></div>';
    const b=k.querySelector('.ob');
    oppg.delopgaver.forEach(d=>mkFI(b,d));
    s.appendChild(k);
  });
});}

if(D.grammatikkData){const gd=D.grammatikkData;aS('s-g','\\ud83d\\udcd8','Grammatikk: '+gd.tema,s=>{
  s.innerHTML='<h2 class="st"><span class="badge">\\ud83d\\udcd8</span> '+e(gd.tema)+'</h2>';
  s.insertAdjacentHTML('beforeend','<div class="gi"><h4>\\ud83d\\udcd8 Forklaring</h4><p>'+e(gd.forklaring)+'</p></div>');
  gd.oppgaver.forEach(oppg=>{
    const ik={fyll_inn:'\\u270f\\ufe0f',multiple_choice:'\\u2611\\ufe0f',ordstilling:'\\ud83d\\udd00',matching:'\\ud83d\\udd17',korriger:'\\ud83d\\udd0d'}[oppg.type]||'\\ud83d\\udcdd';
    const k=document.createElement('div');k.className='ok0';
    k.innerHTML='<div class="oh g">'+ik+' G'+oppg.nummer+': '+e(oppg.tittel)+'</div><div class="ob"><div class="oi">'+e(oppg.instruksjon)+'</div></div>';
    const b=k.querySelector('.ob');
    if(oppg.type==='multiple_choice')oppg.delopgaver.forEach(d=>mkMC(b,d));
    else if(oppg.type==='ordstilling')oppg.delopgaver.forEach(d=>mkOS(b,d));
    else if(oppg.type==='matching')mkMA(b,oppg);
    else oppg.delopgaver.forEach(d=>mkFI(b,d));
    s.appendChild(k);
  });
});}

aS('s-o','\\ud83d\\udcda','Ordliste',s=>{
  s.innerHTML='<h2 class="st"><span class="badge">\\ud83d\\udcda</span> Viktige ord og uttrykk</h2><p class="si">Ordene du har m\\u00f8tt i tekstene.</p><div class="og"></div>';
  const g=s.querySelector('.og');
  D.ordliste.forEach(o=>{
    const k=document.createElement('div');k.className='oc';
    k.innerHTML='<div class="on">'+e(o.norsk)+'</div><div class="of">'+e(o.forklaring||'')+'</div>'+(o.oversettelse?'<div class="oo">'+e(D.hjelpesprak||'')+': '+e(o.oversettelse)+'</div>':'');
    g.appendChild(k);
  });
});

if(sd.length>0)vS(sd[0]);
uS();
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
