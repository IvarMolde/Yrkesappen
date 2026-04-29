'use strict';

const express = require('express');
const rateLimit = require('express-rate-limit');
const { registerAuthRoute } = require('./src/routes/auth');
const { registerGenerateRoute } = require('./src/routes/generate');
const { errorHandler } = require('./src/middleware/error-handler');
const {
  normaliserYrke,
  callGemini,
  buildPrompt,
  buildGrammatikkPrompt,
} = require('./src/services/ai');
const {
  buildDocx,
  buildPptx,
  buildInteractiveHtml,
} = require('./src/services/document-generation');

const app = express();
app.use(express.json({ limit: '10mb' }));
app.use(express.static(__dirname));

const apiLimiter = rateLimit({
  windowMs: 15 * 60 * 1000,
  max: 100,
  standardHeaders: true,
  legacyHeaders: false,
  message: { feil: 'For mange forespørsler. Vent litt og prøv igjen.' },
});

const authLimiter = rateLimit({
  windowMs: 10 * 60 * 1000,
  max: 20,
  standardHeaders: true,
  legacyHeaders: false,
  message: { ok: false, feil: 'For mange innloggingsforsøk. Prøv igjen senere.' },
});

app.use('/api', apiLimiter);
app.use('/api/logginn', authLimiter);

app.get('/api/health', (req, res) => {
  res.json({
    ok: true,
    status: 'healthy',
    uptimeSec: Math.floor(process.uptime()),
    timestamp: new Date().toISOString(),
  });
});

registerAuthRoute(app);
registerGenerateRoute(app, {
  normaliserYrke,
  callGemini,
  buildPrompt,
  buildGrammatikkPrompt,
  buildDocx,
  buildPptx,
  buildInteractiveHtml,
});

app.use(errorHandler);

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Yrkesappen kjører på port ${PORT}`));
