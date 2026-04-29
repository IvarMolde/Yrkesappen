'use strict';
const { loginSchema } = require('../schemas/api');
const { issueAuthToken } = require('../services/auth-token');

function registerAuthRoute(app) {
  app.post('/api/logginn', (req, res) => {
    const parsed = loginSchema.safeParse(req.body);
    if (!parsed.success) {
      return res.status(400).json({ ok: false, feil: 'Ugyldig innlogging-data.' });
    }
    const { passord } = parsed.data;
    const riktig = process.env.APP_PASSORD;
    if (!riktig) return res.status(500).json({ ok: false, feil: 'APP_PASSORD ikke satt.' });
    if (passord === riktig) return res.json({ ok: true, authToken: issueAuthToken() });
    return res.status(401).json({ ok: false });
  });
}

module.exports = { registerAuthRoute };
