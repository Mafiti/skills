#!/usr/bin/env node

import { spawnSync } from 'node:child_process';
import { createRequire } from 'node:module';
import fs from 'node:fs';
import path from 'node:path';
import process from 'node:process';
import { fileURLToPath } from 'node:url';

const SKILL_ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const PACKAGE_JSON_PATH = path.join(SKILL_ROOT, 'package.json');
const PACKAGE_LOCK_PATH = path.join(SKILL_ROOT, 'package-lock.json');
const RUNTIME_DEPENDENCIES = ['exceljs', 'jszip'];
const skillRequire = createRequire(PACKAGE_JSON_PATH);

function dependenciesReady() {
  return RUNTIME_DEPENDENCIES.every((dependencyName) => {
    try {
      skillRequire.resolve(dependencyName);
      return true;
    } catch {
      return false;
    }
  });
}

function installDependencies() {
  const npmCommand = process.platform === 'win32' ? 'npm.cmd' : 'npm';
  const installArgs = fs.existsSync(PACKAGE_LOCK_PATH)
    ? ['ci', '--omit=dev', '--ignore-scripts']
    : ['install', '--omit=dev', '--ignore-scripts'];
  const result = spawnSync(npmCommand, installArgs, {
    cwd: SKILL_ROOT,
    stdio: 'inherit',
  });

  if (result.error) {
    throw new Error(`Unable to start ${npmCommand}: ${result.error.message}`);
  }
  if (result.status !== 0) {
    throw new Error(`Runtime dependency installation failed with exit code ${result.status}.`);
  }
}

if (dependenciesReady()) {
  console.log('HZero release i18n agent runtime dependencies are ready.');
} else {
  console.log('Installing HZero release i18n agent runtime dependencies...');
  installDependencies();
  if (!dependenciesReady()) {
    throw new Error(`Dependencies are still unavailable after installation: ${RUNTIME_DEPENDENCIES.join(', ')}.`);
  }
  console.log('HZero release i18n agent runtime dependencies are ready.');
}
