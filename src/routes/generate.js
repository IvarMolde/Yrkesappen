'use strict';
const { generateSchema } = require('../schemas/api');

function rettForbokstav(tekst, y) {
  if (!tekst || !y) return tekst;
  const stor = y.charAt(0).toUpperCase() + y.slice(1);
  const liten = y.charAt(0).toLowerCase() + y.slice(1);
  return tekst.replace(
    new RegExp(`(?<![.!?]\\s)(?<!^)\\b${stor}(en|er|ene|ens|s)?\\b`, 'g'),
    (match, ending) => liten + (ending || '')
  );
}

function registerGenerateRoute(app, deps) {
  const {
    normaliserYrke,
    callGemini,
    buildPrompt,
    buildGrammatikkPrompt,
    buildDocx,
    buildPptx,
  } = deps;

  app.post('/api/generer', async (req, res, next) => {
    try {
      const parsed = generateSchema.safeParse(req.body);
      if (!parsed.success) {
        return res.status(400).json({ feil: 'Ugyldig input. Sjekk yrke, nivå og valg.' });
      }
      const { yrke, niva, sprak, plassering, fokus, grammatikkFokus, passord } = parsed.data;

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
      if (data.seksjoner) {
        data.seksjoner = data.seksjoner.map((s) => {
          if (s.innhold) s.innhold = rettForbokstav(s.innhold, yrkeNormalisert);
          if (s.tittel) s.tittel = rettForbokstav(s.tittel, yrkeNormalisert);
          if (s.instruksjon) s.instruksjon = rettForbokstav(s.instruksjon, yrkeNormalisert);
          if (s.delopgaver) {
            s.delopgaver = s.delopgaver.map((d) => ({
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
      return next(err);
    }
  });
}

module.exports = { registerGenerateRoute };
