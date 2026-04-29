'use strict';

const crypto = require('crypto');

const TOKEN_TTL_SEC = 8 * 60 * 60;

function base64UrlEncode(input) {
  return Buffer.from(input).toString('base64url');
}

function base64UrlDecode(input) {
  return Buffer.from(input, 'base64url').toString('utf8');
}

function tokenSecret() {
  return process.env.APP_SESSION_SECRET || process.env.APP_PASSORD || 'dev-fallback-secret';
}

function sign(data) {
  return crypto
    .createHmac('sha256', tokenSecret())
    .update(data)
    .digest('base64url');
}

function issueAuthToken() {
  const payload = {
    exp: Math.floor(Date.now() / 1000) + TOKEN_TTL_SEC,
    nonce: crypto.randomBytes(12).toString('hex'),
  };
  const payloadEncoded = base64UrlEncode(JSON.stringify(payload));
  const signature = sign(payloadEncoded);
  return `${payloadEncoded}.${signature}`;
}

function verifyAuthToken(token) {
  if (!token || typeof token !== 'string' || !token.includes('.')) {
    return false;
  }
  const [payloadEncoded, signature] = token.split('.');
  if (!payloadEncoded || !signature) {
    return false;
  }

  const expected = sign(payloadEncoded);
  const a = Buffer.from(signature);
  const b = Buffer.from(expected);
  if (a.length !== b.length || !crypto.timingSafeEqual(a, b)) {
    return false;
  }

  try {
    const payload = JSON.parse(base64UrlDecode(payloadEncoded));
    if (!payload.exp || typeof payload.exp !== 'number') {
      return false;
    }
    return payload.exp > Math.floor(Date.now() / 1000);
  } catch {
    return false;
  }
}

module.exports = {
  issueAuthToken,
  verifyAuthToken,
};
