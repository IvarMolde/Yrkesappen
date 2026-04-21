'use strict';
const express = require('express');
const https = require('https');
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

// ─── Gemini 2.5 Flash via direkte HTTP ────────────────────────────────────────
function callGemini(prompt) {
  return new Promise((resolve, reject) => {
    const apiKey = process.env.GEMINI_API_KEY;
    if (!apiKey) return reject(new Error('GEMINI_API_KEY mangler i miljøvariabler.'));

    const body = JSON.stringify({
      contents: [{ parts: [{ text: prompt }] }],
      // 16384 gir nok rom for B1/B2 + grammatikk + struktur uten å kutte av
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

          // Kritisk: sjekk finishReason – MAX_TOKENS betyr at svaret ble kuttet av
          if (candidate.finishReason && candidate.finishReason !== 'STOP') {
            console.error(`⚠️ Gemini finishReason: ${candidate.finishReason} – svaret kan være ufullstendig`);
          }

          const text = candidate.content && candidate.content.parts && candidate.content.parts[0] && candidate.content.parts[0].text;
          if (!text) return reject(new Error('Ingen tekst i Gemini-svar'));
          resolve(text);
        } catch (e) {
          reject(new Error('Kunne ikke tolke svar fra Gemini: ' + data.slice(0, 200)));
        }
      });
    });
    // Timeout: maks 45 sekunder for Gemini-kall (Vercel Pro gir 60s, gratis 10s)
    req.setTimeout(45000, () => {
      req.destroy();
      reject(new Error('Gemini-kall timet ut etter 45 sekunder'));
    });
    req.on('error', reject);
    req.write(body);
    req.end();
  });
}

// ─── Prompt ────────────────────────────────────────────────────────────────────
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
    {
      "type": "tekst",
      "nummer": 1,
      "tittel": "Tekst 1 – [konkret tema, f.eks. En vanlig arbeidsdag]",
      "innhold": "Fagtekst 1 på ${niva}-nivå. Handle om en konkret hverdagssituasjon. Bruk gjerne dialog eller fortellende form."
    },
    {
      "type": "oppgave",
      "nummer": 1,
      "tilknyttet_tekst": "Tekst 1",
      "oppgavetype": "leseforståelse",
      "tittel": "Leseforståelse – Tekst 1",
      "instruksjon": "Les Tekst 1 og svar på spørsmålene.",
      "delopgaver": [
        { "bokstav": "a", "tekst": "Spørsmål direkte fra Tekst 1..." },
        { "bokstav": "b", "tekst": "..." },
        { "bokstav": "c", "tekst": "..." },
        { "bokstav": "d", "tekst": "..." },
        { "bokstav": "e", "tekst": "..." }
      ]
    },
    {
      "type": "oppgave",
      "nummer": 2,
      "tilknyttet_tekst": "Tekst 1",
      "oppgavetype": "grammatikk",
      "tittel": "Grammatikk – [${grammatikkMap[niva].split(',')[0]}]",
      "instruksjon": "Grammatikkoppgave med eksempler fra Tekst 1.",
      "delopgaver": [
        { "bokstav": "a", "tekst": "..." },
        { "bokstav": "b", "tekst": "..." },
        { "bokstav": "c", "tekst": "..." },
        { "bokstav": "d", "tekst": "..." },
        { "bokstav": "e", "tekst": "..." }
      ]
    },
    {
      "type": "tekst",
      "nummer": 2,
      "tittel": "Tekst 2 – [tema: utstyr, verktøy eller faglige metoder]",
      "innhold": "Fagtekst 2 på ${niva}-nivå om utstyr, verktøy, fagbegreper eller arbeidsmetoder. Litt mer krevende enn Tekst 1."
    },
    {
      "type": "oppgave",
      "nummer": 3,
      "tilknyttet_tekst": "Tekst 2",
      "oppgavetype": "leseforståelse",
      "tittel": "Leseforståelse – Tekst 2",
      "instruksjon": "Les Tekst 2 og svar på spørsmålene.",
      "delopgaver": [
        { "bokstav": "a", "tekst": "..." },
        { "bokstav": "b", "tekst": "..." },
        { "bokstav": "c", "tekst": "..." },
        { "bokstav": "d", "tekst": "..." },
        { "bokstav": "e", "tekst": "..." }
      ]
    },
    {
      "type": "oppgave",
      "nummer": 4,
      "tilknyttet_tekst": "Tekst 2",
      "oppgavetype": "vokabular",
      "tittel": "Ord og uttrykk fra Tekst 2",
      "instruksjon": "Vokabularoppgave basert på ord fra Tekst 2.",
      "delopgaver": [
        { "bokstav": "a", "tekst": "..." },
        { "bokstav": "b", "tekst": "..." },
        { "bokstav": "c", "tekst": "..." },
        { "bokstav": "d", "tekst": "..." },
        { "bokstav": "e", "tekst": "..." }
      ]
    },
    {
      "type": "tekst",
      "nummer": 3,
      "tittel": "Tekst 3 – [tema: samarbeid, kommunikasjon eller HMS]",
      "innhold": "Fagtekst 3 på ${niva}-nivå om samarbeid, kommunikasjon eller HMS. Den mest krevende av de tre tekstene."
    },
    {
      "type": "oppgave",
      "nummer": 5,
      "tilknyttet_tekst": "Tekst 3",
      "oppgavetype": "leseforståelse",
      "tittel": "Leseforståelse – Tekst 3",
      "instruksjon": "Les Tekst 3 og svar på spørsmålene.",
      "delopgaver": [
        { "bokstav": "a", "tekst": "..." },
        { "bokstav": "b", "tekst": "..." },
        { "bokstav": "c", "tekst": "..." },
        { "bokstav": "d", "tekst": "..." },
        { "bokstav": "e", "tekst": "..." }
      ]
    },
    {
      "type": "oppgave",
      "nummer": 6,
      "tilknyttet_tekst": "Generell",
      "oppgavetype": "grammatikk",
      "tittel": "Grammatikk – [${grammatikkMap[niva].split(',')[1] || grammatikkMap[niva].split(',')[0]}]",
      "instruksjon": "Grammatikkoppgave med eksempler fra alle tekstene.",
      "delopgaver": [
        { "bokstav": "a", "tekst": "..." },
        { "bokstav": "b", "tekst": "..." },
        { "bokstav": "c", "tekst": "..." },
        { "bokstav": "d", "tekst": "..." },
        { "bokstav": "e", "tekst": "..." }
      ]
    },
    {
      "type": "oppgave",
      "nummer": 7,
      "tilknyttet_tekst": "Generell",
      "oppgavetype": "vokabular",
      "tittel": "Ord fra alle tekstene",
      "instruksjon": "Vokabularoppgave med ord fra alle tre tekstene.",
      "delopgaver": [
        { "bokstav": "a", "tekst": "..." },
        { "bokstav": "b", "tekst": "..." },
        { "bokstav": "c", "tekst": "..." },
        { "bokstav": "d", "tekst": "..." },
        { "bokstav": "e", "tekst": "..." }
      ]
    },
    {
      "type": "oppgave",
      "nummer": 8,
      "tilknyttet_tekst": "Generell",
      "oppgavetype": "skriv_muntlig",
      "tittel": "Skriv og snakk",
      "instruksjon": "Produksjonsoppgave – skriv og/eller snakk om yrket.",
      "delopgaver": [
        { "bokstav": "a", "tekst": "..." },
        { "bokstav": "b", "tekst": "..." },
        { "bokstav": "c", "tekst": "..." },
        { "bokstav": "d", "tekst": "..." },
        { "bokstav": "e", "tekst": "..." }
      ]
    }
  ],
  "ordliste": [
    { "norsk": "fagord", "forklaring": "norsk forklaring på ${niva}-nivå"${sprak && sprak !== 'ingen' ? `, "oversettelse": "oversettelse på ${sprak} – KUN ${sprak}"` : ''} }
  ],
  "pptx": {
    "nokkelord": ["8 viktige fagord for yrket"],
    "hms": ["4 viktige HMS-punkter"],
    "egenskaper": ["5 personlige egenskaper"],
    "arbeidsoppgaver": ["4 typiske arbeidsoppgaver"],
    "utdanning": ["3 utdanningsveier eller karrieremuligheter"]
  }
}

STRENGE KRAV:
- Fagtekstene MÅ være i den lengden angitt for ${niva} – ikke kortere!
- Tekstene skal bli litt mer krevende fra Tekst 1 til Tekst 3
- Grammatikkoppgavene MÅ bruke eksempler direkte fra fagtekstene
- Grammatikkfokus MÅ stemme med ${niva}: ${grammatikkMap[niva]}
- Alle oppgaver har nøyaktig 5 delopgaver (a–e), der a–c er litt lettere enn d–e
- Ordlisten: 12–16 ord hentet fra alle tre tekstene
- Legg til totalt 10–12 seksjoner i "seksjoner"-arrayet (3 tekster + 7–9 oppgaver)
- NORSK RETTSKRIVING: Yrkestitler skrives ALLTID med liten forbokstav på norsk. Skriv «sykepleier», ikke «Sykepleier». Skriv «begravelsesagent», ikke «Begravelsesagent». Dette gjelder inne i setninger, i oppgavetitler og overalt i teksten. Unntaket er kun om yrket starter en setning.

ORDLISTE-FORMAT (KRITISK – følg nøyaktig):
- SUBSTANTIV: ALLTID vis med ubestemt artikkel foran. Bruk «en» for hankjønn, «ei» eller «en» for hunkjønn (bokmål tillater begge), «et» for intetkjønn. Eksempler:
  • «en pasient» (ikke «pasient»)
  • «et sykehus» (ikke «sykehus»)
  • «en journal» (ikke «journal»)
  • «et verneutstyr» (ikke «verneutstyr»)
- VERB: ALLTID vis i infinitiv med infinitivsmerket «å» foran. Eksempler:
  • «å behandle» (ikke «behandle», ikke «behandler»)
  • «å undersøke» (ikke «undersøke»)
  • «å starte» (ikke «starter»)
- ADJEKTIV: vis i grunnform (ubestemt hankjønn entall), f.eks. «rask», «profesjonell», «ansvarlig»
- UTTRYKK/FLERE ORD: behold naturlig form, f.eks. «å ta ansvar», «førstehjelp»
- Dette formatet er ABSOLUTT – alle ord i "norsk"-feltet MÅ følge disse reglene${sprak && sprak !== 'ingen' ? `\n- OVERSETTELSE: Hvert "oversettelse"-felt MÅ inneholde ${sprak}. KUN ${sprak}. Kontroller hvert felt før du svarer.` : ''}`;
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

  return `Du er en erfaren norsklærer som lager grammatikkmateriell for voksne innvandrere (norsk som andrespråk) på CEFR-nivå ${niva}.
Yrke for eksempler og setninger: "${yrke}".

${fokusInstruksjon}

NIVÅTILPASNING ${niva}: ${nivaBeskriv}

LAG EN GRAMMATIKKBLOKK med forklaring og 5 oppgaver (se JSON-struktur under).

KRAV TIL FORKLARINGEN:
- Pedagogisk og faglig korrekt norsk bokmål
- Forklar: hva regelen er, når den brukes, og gi 2-3 korte eksempler fra arbeidslivet
- Tilpasset ${niva}-nivå – enkelt språk for lavere nivå, mer presist for høyere
- Maks 120 ord

KRAV TIL OPPGAVENE:
- Alle setninger og eksempler handler om yrket "${yrke}" eller arbeidslivet
- Stigende vanskelighetsgrad innen hver oppgave (a enklest → e vanskeligst)
- Norsk rettskriving: yrkestitler med liten forbokstav inne i setninger
- Alle "fasit"-felt MÅ inneholde 100% korrekt norsk bokmål
- Oppgave 1 (fyll_inn): Bruk parentes rundt ord som skal bøyes, f.eks. "(jobbe) → jobbet"
- Oppgave 2 (multiple_choice): Alltid nøyaktig 3 alternativer, kun 1 korrekt
- Oppgave 3 (ordstilling): Skill ord med " / ", f.eks. "i dag / jobber / hun / tidlig"
- Oppgave 4 (matching): Nøyaktig 5 par (a-e), kolonne A og kolonne B
- Oppgave 5 (korriger): Én og bare én grammatisk feil per setning

Svar KUN med gyldig JSON, ingen markdown:

{
  "tema": "Kort navn på grammatikktemaet (maks 5 ord)",
  "forklaring": "Pedagogisk forklaring tilpasset ${niva}. Inkluder hva regelen er, når den brukes, og 2-3 eksempler fra arbeidslivet til ${yrke}.",
  "oppgaver": [
    {
      "nummer": 1,
      "type": "fyll_inn",
      "tittel": "Fyll inn riktig form",
      "instruksjon": "Fyll inn riktig form av ordet i parentes.",
      "delopgaver": [
        { "bokstav": "a", "tekst": "Setning med (infinitiv) som skal bøyes.", "fasit": "korrekt bøyd form" },
        { "bokstav": "b", "tekst": "...", "fasit": "..." },
        { "bokstav": "c", "tekst": "...", "fasit": "..." },
        { "bokstav": "d", "tekst": "...", "fasit": "..." },
        { "bokstav": "e", "tekst": "...", "fasit": "..." }
      ]
    },
    {
      "nummer": 2,
      "type": "multiple_choice",
      "tittel": "Velg riktig alternativ",
      "instruksjon": "Velg det riktige alternativet (a, b eller c).",
      "delopgaver": [
        { "bokstav": "a", "tekst": "Setning med ___ der ett ord mangler.", "alternativer": ["alt1", "alt2", "alt3"], "fasit": "riktig alternativ" },
        { "bokstav": "b", "tekst": "...", "alternativer": ["...", "...", "..."], "fasit": "..." },
        { "bokstav": "c", "tekst": "...", "alternativer": ["...", "...", "..."], "fasit": "..." },
        { "bokstav": "d", "tekst": "...", "alternativer": ["...", "...", "..."], "fasit": "..." },
        { "bokstav": "e", "tekst": "...", "alternativer": ["...", "...", "..."], "fasit": "..." }
      ]
    },
    {
      "nummer": 3,
      "type": "ordstilling",
      "tittel": "Sett ordene i riktig rekkefølge",
      "instruksjon": "Sett ordene i riktig rekkefølge og skriv hele setningen.",
      "delopgaver": [
        { "bokstav": "a", "tekst": "ord1 / ord2 / ord3 / ord4", "fasit": "Riktig setning." },
        { "bokstav": "b", "tekst": "...", "fasit": "..." },
        { "bokstav": "c", "tekst": "...", "fasit": "..." },
        { "bokstav": "d", "tekst": "...", "fasit": "..." },
        { "bokstav": "e", "tekst": "...", "fasit": "..." }
      ]
    },
    {
      "nummer": 4,
      "type": "matching",
      "tittel": "Koble sammen",
      "instruksjon": "Koble hvert uttrykk i kolonne A med riktig svar i kolonne B.",
      "delopgaver": [
        { "bokstav": "a", "tekst": "Kolonne A: setningsdel eller ord", "match": "Kolonne B: riktig fortsettelse" },
        { "bokstav": "b", "tekst": "...", "match": "..." },
        { "bokstav": "c", "tekst": "...", "match": "..." },
        { "bokstav": "d", "tekst": "...", "match": "..." },
        { "bokstav": "e", "tekst": "...", "match": "..." }
      ]
    },
    {
      "nummer": 5,
      "type": "korriger",
      "tittel": "Korriger feilen",
      "instruksjon": "Hver setning inneholder én grammatisk feil. Skriv setningen riktig.",
      "delopgaver": [
        { "bokstav": "a", "tekst": "Setning med én grammatisk feil.", "fasit": "Riktig setning." },
        { "bokstav": "b", "tekst": "...", "fasit": "..." },
        { "bokstav": "c", "tekst": "...", "fasit": "..." },
        { "bokstav": "d", "tekst": "...", "fasit": "..." },
        { "bokstav": "e", "tekst": "...", "fasit": "..." }
      ]
    }
  ]
}`;
}



// ─── Normalisér yrkenavn via Gemini ───────────────────────────────────────────
async function normaliserYrke(yrke) {
  const apiKey = process.env.GEMINI_API_KEY;
  const input = yrke.toLowerCase().trim();
  if (!apiKey) return { yrke: input, korrigert: false };

  const prompt = `Du er en norsk rettskrivingsekspert. En bruker har skrevet inn et yrkenavn på norsk bokmål.

Brukerens input: "${input}"

Gjør følgende:
1. Korriger eventuelle stavefeil til korrekt norsk bokmål
2. Yrkestitler skal ALLTID ha liten forbokstav (f.eks. «sykepleier», ikke «Sykepleier»)
3. Hvis input allerede er korrekt, returner det uendret

Svar KUN med JSON, ingen annen tekst:
{ "yrke": "korrigert yrkenavn", "endret": true/false }

Eksempler:
"Sykepleier" → { "yrke": "sykepleier", "endret": true }
"sykeplier" → { "yrke": "sykepleier", "endret": true }
"elektrikar" → { "yrke": "elektriker", "endret": true }
"kokk" → { "yrke": "kokk", "endret": false }
"begraffelsesagent" → { "yrke": "begravelsesagent", "endret": true }`;

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
          console.log(`Yrke: "${input}" → "${korrigert}" (endret: ${result.endret})`);
          resolve({ yrke: korrigert, korrigert: result.endret === true && korrigert !== input });
        } catch (e) {
          console.error('Normalisering feilet:', e.message);
          resolve({ yrke: input, korrigert: false });
        }
      });
    });
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
        keepNext: true,  // Hold sammen med neste paragraf
        keepLines: true,
        children: [new TextRun({ text: `  Oppgave ${nr}  `, bold: true, size: 26, color: C.white, font: 'Calibri' })],
      }),
      ...(visTekst ? [new Paragraph({
        spacing: { before: 40, after: 40 },
        shading: { fill: C.secondary, type: ShadingType.CLEAR },
        keepNext: true,
        keepLines: true,
        children: [new TextRun({ text: `  ${typeIkon} ${tilknyttetTekst}  `, size: 20, color: C.white, font: 'Calibri' })],
      })] : []),
      new Paragraph({
        spacing: { before: 60, after: 60 },
        keepNext: true,
        keepLines: true,
        children: [new TextRun({ text: tittel, bold: true, size: 28, color: C.textDark, font: 'Calibri' })],
      }),
      new Paragraph({
        spacing: { after: 120 },
        keepNext: true,  // Instruksjonen må holdes sammen med tabellen under
        keepLines: true,
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

  // Seksjoner (tekster og oppgaver flettet)
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
          cantSplit: true, // hindrer at raden deles over to sider
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

  // Ordliste
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

  // ── Grammatikkblokk (valgfritt) ──────────────────────────────────────────────
  const grammatikkBlock = [];
  if (grammatikkData && grammatikkData.oppgaver) {
    const typeIkonGram = {
      fyll_inn:       '✏️',
      multiple_choice:'☑️',
      ordstilling:    '🔀',
      matching:       '🔗',
      korriger:       '🔍',
    };

    // Seksjonstittel
    grammatikkBlock.push(...sectionHeader(`Grammatikk: ${grammatikkData.tema}`));

    // Forklaringsboks med teal bakgrunn
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

    // Fem oppgaver
    grammatikkData.oppgaver.forEach((oppg) => {
      const ikon = typeIkonGram[oppg.type] || '📝';

      // Oppgaveheader
      grammatikkBlock.push(
        new Paragraph({
          spacing: { before: 280, after: 60 },
          shading: { fill: C.secondary, type: ShadingType.CLEAR },
          keepNext: true,
          keepLines: true,
          children: [new TextRun({ text: `  ${ikon} Oppgave G${oppg.nummer}: ${oppg.tittel}  `, bold: true, size: 24, color: C.white, font: 'Calibri' })],
        }),
        new Paragraph({
          spacing: { before: 60, after: 100 },
          keepNext: true,
          keepLines: true,
          children: [new TextRun({ text: oppg.instruksjon, italics: true, size: 22, color: C.textMid, font: 'Calibri' })],
        })
      );

      // Delopgaver – ulik layout per type
      oppg.delopgaver.forEach((d, idx) => {
        const fill = idx % 2 === 0 ? C.white : C.bgGray;

        if (oppg.type === 'multiple_choice' && d.alternativer) {
          // Multiple choice: vis alternativer på egen linje
          grammatikkBlock.push(
            new Paragraph({
              spacing: { before: 60, after: 20 },
              shading: { fill, type: ShadingType.CLEAR },
              keepNext: true,
              keepLines: true,
              children: [
                new TextRun({ text: `${d.bokstav})  `, bold: true, size: 24, color: C.primary, font: 'Calibri' }),
                new TextRun({ text: d.tekst, size: 24, font: 'Calibri' }),
              ],
            }),
            new Paragraph({
              spacing: { before: 0, after: 20 },
              shading: { fill, type: ShadingType.CLEAR },
              indent: { left: 360 },
              keepNext: true,
              keepLines: true,
              children: d.alternativer.map((alt, ai) =>
                new TextRun({ text: `  ${['A', 'B', 'C'][ai]}) ${alt}   `, size: 22, font: 'Calibri', color: C.textMid })
              ),
            }),
            svarLinje()
          );
        } else if (oppg.type === 'matching') {
          // Matching: to kolonner – cantSplit hindrer deling av rad
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
          // fyll_inn, ordstilling, korriger: standard to-kolonne layout – cantSplit hindrer deling
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

      // Fasit-seksjon (sammenleggbar visuelt – grå bakgrunn)
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

// ─── PPTX builder ──────────────────────────────────────────────────────────────
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
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0,     w: 10, h: 0.2, fill: { color: C.accent }, line: { color: C.accent } });
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

  // ── Slide 1 – Tittelslide ────────────────────────────────────────────────────
  {
    const s = pres.addSlide();
    s.background = { color: C.primary };
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0,     w: 10, h: 0.2, fill: { color: C.accent }, line: { color: C.accent } });
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.425, w: 10, h: 0.2, fill: { color: C.accent }, line: { color: C.accent } });
    s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.2, w: 9, h: 2.5, fill: { color: '000000', transparency: 45 }, line: { color: '000000', transparency: 45 } });
    s.addText(yrke.toUpperCase(), { x: 0.5, y: 1.3, w: 9, h: 1.5, fontSize: 44, bold: true, color: C.white, fontFace: 'Calibri', align: 'center', valign: 'middle', wrap: true, shrinkText: true, margin: 8 });
    s.addText(`Norsknivå ${niva}`, { x: 0.5, y: 2.85, w: 9, h: 0.5, fontSize: 20, color: C.accent, fontFace: 'Calibri', align: 'center', valign: 'middle', margin: 0 });
    s.addText('Molde voksenopplæringssenter – MBO', { x: 0.5, y: 3.95, w: 9, h: 0.4, fontSize: 13, color: C.bgGray, fontFace: 'Calibri', align: 'center', valign: 'middle', margin: 0 });
    if (hasFokus) s.addText(`Fokus: ${fokus}`, { x: 1.5, y: 4.1, w: 7, h: 0.5, fontSize: 13, italic: true, color: C.accent, fontFace: 'Calibri', align: 'center', valign: 'middle', wrap: true, shrinkText: true, margin: 4 });
  }

  // ── Slide 2 – Hva er dette yrket? ───────────────────────────────────────────
  {
    const s = lightSlide('Hva er dette yrket?', 'Forberedelse til arbeidsheftet');
    const items = arbeidsoppgaver.map((t, idx) => ({ text: t, options: { bullet: true, breakLine: idx < arbeidsoppgaver.length - 1, fontSize: 15, color: C.textDark, fontFace: 'Calibri', paraSpaceAfter: 8 } }));
    s.addText(items, { x: 0.25, y: 1.1, w: 5.8, h: 4.3, valign: 'top', wrap: true, shrinkText: true, margin: 8 });
    // Høyre panel: stor bokstav på teal bakgrunn
    s.addShape(pres.shapes.RECTANGLE, { x: 6.3, y: 1.1, w: 3.45, h: 4.3, fill: { color: C.primary }, line: { color: C.primary }, shadow: makeShadow() });
    s.addText(yrke.charAt(0).toUpperCase(), { x: 6.3, y: 1.1, w: 3.45, h: 3.0, fontSize: 110, bold: true, color: C.white, fontFace: 'Calibri', align: 'center', valign: 'middle', margin: 0 });
    safeText(s, yrke, 6.35, 4.1, 3.35, 0.7, { bold: true, color: C.accent, fontSize: 14, align: 'center', valign: 'middle', margin: 4 });
  }

  // ── Slide 3 – Viktige ord ────────────────────────────────────────────────────
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

  // ── Slides 4, 5, 6 – Forberedelse til fagtekstene ───────────────────────────
  tekster.forEach((tekst, i) => {
    const s = lightSlide(tekst.tittel, `Tekst ${tekst.nummer} av 3 – forkunnskaper og forberedelse`);
    const aktiviteter = [
      { ikon: '🤔', label: 'Tenk – Par – Del', sporsmal: `Hva tror du en ${yrke} gjør i løpet av en arbeidsdag?`, tips: ['Tenk selv i 30 sekunder', 'Snakk med sidemannen din', 'Del svaret med klassen'] },
      { ikon: '💡', label: 'Brainstorm',        sporsmal: `Hvilket utstyr eller hvilke verktøy tror du en ${yrke} bruker?`, tips: ['Skriv ned så mange ord du kan', 'Sammenlign med sidemannen', 'Hvilke ord kjenner du fra før?'] },
      { ikon: '🤝', label: 'Del erfaring',      sporsmal: 'Hva er viktig når man jobber sammen med andre mennesker?', tips: ['Tenk på din egen arbeidserfaring', 'Hva er god kommunikasjon på jobb?', 'Hva skjer hvis sikkerhetsregler ikke følges?'] },
    ][i] || { ikon: '💬', label: 'Diskuter', sporsmal: `Hva vet du om yrket ${yrke} fra før?`, tips: ['Snakk med sidemannen', 'Del med klassen'] };

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

  // ── Slide 7 – HMS ────────────────────────────────────────────────────────────
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

  // ── Slide 8 – Personlige egenskaper ─────────────────────────────────────────
  {
    const s = lightSlide('Personlige egenskaper', 'Hva er viktig for å lykkes i dette yrket?');
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

  // ── Slide 9 – Utdanning og karriere ─────────────────────────────────────────
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

  // ── Slide 10 – La oss snakke norsk! ─────────────────────────────────────────
  {
    const s = lightSlide('La oss snakke norsk! 💬', 'Diskuter med sidepersonen din');
    const sporsmal = [
      `Hva vet du om yrket ${yrke}?`,
      'Ville du likt å jobbe i dette yrket? Hvorfor / hvorfor ikke?',
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

// ─── Innlogging ────────────────────────────────────────────────────────────────
app.post('/api/logginn', (req, res) => {
  const { passord } = req.body;
  const riktig = process.env.APP_PASSORD;
  if (!riktig) return res.status(500).json({ ok: false, feil: 'APP_PASSORD ikke satt.' });
  if (passord === riktig) return res.json({ ok: true });
  return res.status(401).json({ ok: false });
});

// ─── Generer endpoint ──────────────────────────────────────────────────────────
app.post('/api/generer', async (req, res) => {
  try {
    const { yrke, niva, sprak, plassering, fokus, grammatikkFokus, passord } = req.body;
    if (!yrke || !niva) return res.status(400).json({ feil: 'Yrke og nivå er påkrevd.' });

    const riktig = process.env.APP_PASSORD;
    if (riktig && passord !== riktig) return res.status(401).json({ feil: 'Ikke autorisert. Logg inn på nytt.' });

    // Normalisér og korriger yrkenavnet (stavefeil + liten forbokstav)
    const { yrke: yrkeNormalisert, korrigert: yrkeKorrigert } = await normaliserYrke(yrke);
    console.log(`Yrke brukt i generering: "${yrkeNormalisert}"`);

    const raw = await callGemini(buildPrompt(yrkeNormalisert, niva, sprak, plassering, fokus));
    const clean = raw.trim().replace(/^```(?:json)?\s*/i, '').replace(/\s*```\s*$/i, '');

    let data;
    try {
      data = JSON.parse(clean);
    } catch (e) {
      console.error('JSON feil:', clean.slice(0, 400));
      return res.status(500).json({ feil: 'Klarte ikke tolke svar fra AI. Prøv igjen.' });
    }

    // ── Sikkerhetsnett: rett opp stor forbokstav på yrkestittel ───────────────
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
        if (s.tittel)  s.tittel  = rettForbokstav(s.tittel,  yrkeNormalisert);
        if (s.instruksjon) s.instruksjon = rettForbokstav(s.instruksjon, yrkeNormalisert);
        if (s.delopgaver) {
          s.delopgaver = s.delopgaver.map(d => ({
            ...d,
            tekst: rettForbokstav(d.tekst, yrkeNormalisert),
          }));
        }
        return s;
      });
    }
    if (data.intro) data.intro = rettForbokstav(data.intro, yrkeNormalisert);

    // Generer grammatikkblokk om ønsket
    const gFokus = (grammatikkFokus || 'ingen').trim();
    let grammatikkData = null;
    if (gFokus !== 'ingen') {
      console.log(`Genererer grammatikkblokk: "${gFokus}"`);
      try {
        const gRaw = await callGemini(buildGrammatikkPrompt(yrkeNormalisert, niva, gFokus));
        const gClean = gRaw.trim().replace(/^```(?:json)?\s*/i, '').replace(/\s*```\s*$/i, '');
        grammatikkData = JSON.parse(gClean);
        console.log(`Grammatikkblokk generert: ${grammatikkData.tema}`);
      } catch (e) {
        console.error('Grammatikk-generering feilet:', e.message);
      }
    }

    const [docxBuf, pptxBuf] = await Promise.all([
      buildDocx(data, sprak, plassering, grammatikkData),
      buildPptx(data, yrkeNormalisert, niva, sprak, fokus),
    ]);

    const safeName = yrkeNormalisert.replace(/[^a-zA-ZæøåÆØÅ0-9\-]/g, '_');
    res.json({
      docx: docxBuf.toString('base64'),
      pptx: pptxBuf.toString('base64'),
      filnavn: safeName,
      niva,
      yrkeKorrigert: yrkeKorrigert ? yrkeNormalisert : null,
    });

  } catch (err) {
    console.error(err);
    if (!res.headersSent) res.status(500).json({ feil: err.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Yrkesappen kjører på port ${PORT}`));
