'use strict';

function registerAuthRoute(app) {
  app.post('/api/logginn', (req, res) => {
    const { passord } = req.body;
    const riktig = process.env.APP_PASSORD;
    if (!riktig) return res.status(500).json({ ok: false, feil: 'APP_PASSORD ikke satt.' });
    if (passord === riktig) return res.json({ ok: true });
    return res.status(401).json({ ok: false });
  });
}

module.exports = { registerAuthRoute };
