'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const { spawn } = require('node:child_process');

const PORT = 3310;
const BASE_URL = `http://127.0.0.1:${PORT}`;

let serverProc;

function wait(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function waitForHealthy(maxAttempts = 20) {
  for (let i = 0; i < maxAttempts; i += 1) {
    try {
      const res = await fetch(`${BASE_URL}/api/health`);
      if (res.ok) return;
    } catch {
      // waiting for server startup
    }
    await wait(250);
  }
  throw new Error('Server did not become healthy in time.');
}

test.before(async () => {
  serverProc = spawn(process.execPath, ['server.js'], {
    env: {
      ...process.env,
      APP_PASSORD: 'testpass',
      PORT: String(PORT),
    },
    stdio: 'ignore',
  });

  await waitForHealthy();
});

test.after(() => {
  if (serverProc && !serverProc.killed) {
    serverProc.kill();
  }
});

test('health endpoint returns healthy status', async () => {
  const res = await fetch(`${BASE_URL}/api/health`);
  assert.equal(res.status, 200);
  const data = await res.json();
  assert.equal(data.ok, true);
  assert.equal(data.status, 'healthy');
});

test('login returns auth token', async () => {
  const res = await fetch(`${BASE_URL}/api/logginn`, {
    method: 'POST',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify({ passord: 'testpass' }),
  });
  assert.equal(res.status, 200);
  const data = await res.json();
  assert.equal(data.ok, true);
  assert.equal(typeof data.authToken, 'string');
  assert.equal(data.authToken.includes('.'), true);
});

test('generate rejects invalid payload with 400', async () => {
  const res = await fetch(`${BASE_URL}/api/generer`, {
    method: 'POST',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify({ yrke: '', niva: 'C1' }),
  });
  assert.equal(res.status, 400);
});
