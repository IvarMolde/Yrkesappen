'use strict';

const express = require('express');
const { registerAuthRoute } = require('./src/routes/auth');
const { registerGenerateRoute } = require('./src/routes/generate');
const {
  normaliserYrke,
  callGemini,
  buildPrompt,
  buildGrammatikkPrompt,
} = require('./src/services/ai');
const {
  buildDocx,
  buildPptx,
} = require('./src/services/document-generation');

const app = express();
app.use(express.json({ limit: '10mb' }));
app.use(express.static(__dirname));

registerAuthRoute(app);
registerGenerateRoute(app, {
  normaliserYrke,
  callGemini,
  buildPrompt,
  buildGrammatikkPrompt,
  buildDocx,
  buildPptx,
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Yrkesappen kjører på port ${PORT}`));
