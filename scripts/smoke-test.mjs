#!/usr/bin/env node

import fs from 'node:fs';
import path from 'node:path';
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';

const here = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(here, '..');
const input = path.join(root, 'examples', 'demo.html');
const output = path.join(root, 'examples', 'demo-output.pptx');
const report = path.join(root, 'examples', 'demo-report.json');

for (const file of [output, report]) {
  if (fs.existsSync(file)) fs.rmSync(file);
}

const result = spawnSync(process.execPath, [
  path.join(root, 'scripts', 'convert.mjs'),
  input,
  output,
  '--selector', '.slide',
  '--report', report,
], { stdio: 'inherit' });

if (result.status !== 0) process.exit(result.status ?? 1);
if (!fs.existsSync(output) || fs.statSync(output).size < 1024) {
  throw new Error('Smoke test failed: PPTX was not created or is unexpectedly small.');
}
if (!fs.existsSync(report)) {
  throw new Error('Smoke test failed: conversion report was not created.');
}

const parsed = JSON.parse(fs.readFileSync(report, 'utf8'));
if (parsed.totals?.slides !== 1 || (parsed.totals?.text ?? 0) < 1) {
  throw new Error(`Smoke test failed: unexpected report ${JSON.stringify(parsed.totals)}`);
}

console.log(`Smoke test passed: ${output}`);
