'use strict';

const { z } = require('zod');

const MAX_TEXT = 400;

const loginSchema = z.object({
  passord: z.string().trim().min(1).max(200),
});

const generateSchema = z.object({
  yrke: z.string().trim().min(2).max(80),
  niva: z.enum(['A1', 'A2', 'B1', 'B2']),
  sprak: z.string().trim().min(1).max(80).optional().default('ingen'),
  plassering: z.enum(['ingen', 'tosidig', 'slutt']).optional().default('ingen'),
  fokus: z.string().trim().max(MAX_TEXT).optional().default(''),
  grammatikkFokus: z.string().trim().max(MAX_TEXT).optional().default('ingen'),
  passord: z.string().trim().max(200).optional().default(''),
  authToken: z.string().trim().max(800).optional().default(''),
});

module.exports = {
  loginSchema,
  generateSchema,
};
