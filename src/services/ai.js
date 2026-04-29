'use strict';

const https = require('https');

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
      res.on('data', (chunk) => {
        data += chunk;
      });
      res.on('end', () => {
        try {
          const parsed = JSON.parse(data);
          if (parsed.error) return reject(new Error(parsed.error.message));
          const candidate = parsed.candidates && parsed.candidates[0];
          if (!candidate) return reject(new Error('Tomt Gemini-svar'));

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

    req.setTimeout(45000, () => {
      req.destroy();
      reject(new Error('Gemini-kall timet ut etter 45 sekunder'));
    });
    req.on('error', reject);
    req.write(body);
    req.end();
  });
}

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
    { "type": "tekst", "nummer": 1, "tittel": "Tekst 1 – [tema]", "innhold": "Fagtekst 1 på ${niva}-nivå." },
    { "type": "oppgave", "nummer": 1, "tilknyttet_tekst": "Tekst 1", "oppgavetype": "leseforståelse", "tittel": "Leseforståelse – Tekst 1", "instruksjon": "Les Tekst 1 og svar på spørsmålene.", "delopgaver": [{ "bokstav": "a", "tekst": "..." }, { "bokstav": "b", "tekst": "..." }, { "bokstav": "c", "tekst": "..." }, { "bokstav": "d", "tekst": "..." }, { "bokstav": "e", "tekst": "..." }] }
  ],
  "ordliste": [
    { "norsk": "fagord", "forklaring": "norsk forklaring på ${niva}-nivå"${sprak && sprak !== 'ingen' ? `, "oversettelse": "oversettelse på ${sprak} – KUN ${sprak}"` : ''} }
  ],
  "pptx": {
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
- Ordlisten: 12–16 ord hentet fra alle tre tekstene`;
}

function buildGrammatikkPrompt(yrke, niva, grammatikkFokus) {
  const nivaBeskriv = {
    A1: 'Svært enkle setninger. Maks 6-8 ord per setning. Kun presens. Hverdagsord.',
    A2: 'Enkle setninger. Maks 10 ord. Presens og preteritum. Kjente faguttrykk.',
    B1: 'Moderat komplekse setninger. Presens, preteritum, perfektum. Fagord forklares.',
    B2: 'Komplekse setninger. Alle tider, passiv, relativsetninger. Fagord brukes naturlig.',
  }[niva] || '';

  const erTilfeldig = grammatikkFokus === 'tilfeldig';
  const fokusInstruksjon = erTilfeldig
    ? `Velg selv et passende grammatisk tema for nivå ${niva}, basert på CEFR.`
    : `Grammatisk tema: "${grammatikkFokus}". Tilpass forklaringen og ALLE oppgavene nøyaktig til dette temaet på nivå ${niva}.`;

  return `Du er en erfaren norsklærer som lager grammatikkmateriell for voksne innvandrere (norsk som andrespråk) på CEFR-nivå ${niva}.
Yrke for eksempler og setninger: "${yrke}".
${fokusInstruksjon}
NIVÅTILPASNING ${niva}: ${nivaBeskriv}
LAG EN GRAMMATIKKBLOKK med forklaring og 5 oppgaver.
Svar KUN med gyldig JSON:
{
  "tema": "Kort navn på grammatikktemaet (maks 5 ord)",
  "forklaring": "Pedagogisk forklaring",
  "oppgaver": []
}`;
}

async function normaliserYrke(yrke) {
  const apiKey = process.env.GEMINI_API_KEY;
  const input = yrke.toLowerCase().trim();
  if (!apiKey) return { yrke: input, korrigert: false };

  const prompt = `Du er en norsk rettskrivingsekspert. En bruker har skrevet inn et yrkenavn på norsk bokmål.
Brukerens input: "${input}"
Svar KUN med JSON:
{ "yrke": "korrigert yrkenavn", "endret": true/false }`;

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
      res.on('data', (chunk) => {
        data += chunk;
      });
      res.on('end', () => {
        try {
          const parsed = JSON.parse(data);
          const tekst = parsed.candidates[0].content.parts[0].text.trim()
            .replace(/^```(?:json)?\s*/i, '').replace(/\s*```\s*$/i, '');
          const result = JSON.parse(tekst);
          const korrigert = (result.yrke || input).toLowerCase().trim();
          resolve({ yrke: korrigert, korrigert: result.endret === true && korrigert !== input });
        } catch (e) {
          resolve({ yrke: input, korrigert: false });
        }
      });
    });
    req.on('error', () => resolve({ yrke: input, korrigert: false }));
    req.write(body);
    req.end();
  });
}

module.exports = {
  normaliserYrke,
  callGemini,
  buildPrompt,
  buildGrammatikkPrompt,
};
