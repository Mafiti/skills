#!/usr/bin/env node

import { spawnSync } from 'node:child_process';
import { createRequire } from 'node:module';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import process from 'node:process';
import { fileURLToPath } from 'node:url';

const EXCEL_OUTPUTS = ['\u591a\u8bed\u8a00\u5bfc\u51fa.xlsx', '\u79cd\u5b50\u6570\u636e\u5bfc\u51fa.xlsx'];
const SKILL_NAME = 'hzero-release-i18n-agent';
const SKILL_ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const DEFAULT_CONFIG_HOME = process.env.XDG_CONFIG_HOME || path.join(os.homedir(), '.config');
const DEFAULT_STATE_HOME = process.env.XDG_STATE_HOME || path.join(os.homedir(), '.local', 'state');
const DEFAULT_CONFIG_PATH = path.join(DEFAULT_CONFIG_HOME, SKILL_NAME, 'config.json');
const DEFAULT_AGENT_HOME = path.join(DEFAULT_STATE_HOME, SKILL_NAME);
const SCRIPT_PATH_EXAMPLE = '<skill-directory>/scripts/release-i18n-agent.mjs';
const SKILL_REQUIRE = createRequire(path.join(SKILL_ROOT, 'package.json'));
const DEFAULT_DB_BATCH_SIZE = 300;
const DEFAULT_PROFILE_NAME = 'default';
const DEFAULT_RELEASE_BRANCH_PATTERN = 'v{major}.{minor}.x';
const DEFAULT_PUBLISH_COMMAND = 'yarn transpile && npm publish';
const SOURCE_EXTENSIONS = ['.js', '.jsx', '.ts', '.tsx'];
const DETECT_REGEXP = /(intl|IntlUtils)\s*\.\s*(get\s*\(\s*["'`]([\w.-]+)["'`][\s\S]*?\)\s*\.\s*(defaultMessage|d)\s*\(\s*[`"']([\s\S]+?)[`"']|getHTML*\s*\(\s*["'`]([\w.-]+)["'`][\s\S]*?\)\s*\.\s*(defaultMessage|d)\s*\(\s*(["'`<][\s\S]+?["'`>]))[\s*,]*\)/gm;
const NO_DEFAULT_REGEXP = /intl\s*\.\s*get(HTML)*\s*\(\s*["'`]([\w.-]+)["'`]['"\w.,{}:\s-]*\)\s*(?!\s*\.\s*(d|defaultMessage)\s*\(\s*[<'"`\s\\]+)/gm;
const FIXED_TENANT_ID = 0;
const DEFAULT_SEED_WORKSHEET_NAME = '\u591a\u8bed\u8a00\u6570\u636e';
const SEED_WORKBOOK_LAST_UPDATE_VALUE = 'now';
const SEED_WORKBOOK_LAST_UPDATED_BY = '-1';
const SEED_COLUMN_KEYS = {
  promptId: 'prompt_id',
  tenantId: 'tenant_id',
  promptKey: 'prompt_key',
  promptCode: 'prompt_code',
  lang: 'lang',
  description: 'description',
  lastUpdateDate: 'last_update_date',
  lastUpdatedBy: 'last_updated_by',
};
const FIXED_PROMPT_SCHEMA = {
  database: 'hzero_platform',
  table: 'hpfm_prompt',
  columns: {
    promptKey: 'prompt_key',
    promptCode: 'prompt_code',
    lang: 'lang',
    description: 'description',
    tenantId: 'tenant_id',
  },
};

function printHelp() {
  console.log(`Usage:
  node ${SCRIPT_PATH_EXAMPLE} \\
    --target package-name@1.2.3 \\
    [--execute --allow-branch-switch --allow-pull --allow-publish --allow-commit --allow-push]

Manual mode:
  node ${SCRIPT_PATH_EXAMPLE} \\
    --repo-root /path/to/repo \\
    --package-dir packages/pkg \\
    --version 1.2.3 \\
    [--config ${DEFAULT_CONFIG_PATH}] \\
    [--publish-command "yarn transpile && npm publish --tag beta"] \\
    [--execute --allow-branch-switch --allow-pull --allow-publish --allow-commit --allow-push]

Default mode is dry-run. No file is changed and no command is executed unless --execute is set.

Initialize the local database config before db-backed multilingual diff:
  node ${SCRIPT_PATH_EXAMPLE} \\
    --init-db-config \\
    --profile default \\
    --db-type mysql \\
    --db-host 127.0.0.1 \\
    --db-port 3306 \\
    --db-user root \\
    --db-password '<password>'

Required:
  --target <package@version> Shorthand target. Example: package-name@1.9.1-beta.0.
  --project <name>         Project/workspace key in local config.
  --package <name>         Package key under the project config.
  --repo-root <path>       Repository root. Required only without --project.
  --package-dir <path>     Package directory. Required only without configured --package.
  --version <version>      Target package version.

Execution controls:
  --execute                Write version and run allowed commands.
  --allow-branch-switch    Allow checkout to the inferred or configured release branch.
  --allow-pull             Allow git pull --ff-only from the repository root on the release branch.
  --allow-publish          Allow running the publish command.
  --allow-commit           Allow creating a local git commit.
  --allow-push             Allow pushing the release commit after publish succeeds.
  --skip-branch            Do not checkout the release branch.
  --skip-pull              Do not pull the release branch.
  --skip-publish           Treat publish as already done or intentionally skipped.
  --skip-extract           Do not run intl extraction.
  --skip-commit            Do not create a git commit.
  --skip-push              Do not push the release commit.

Command overrides:
  --config <path>          Local database config path. Defaults to ${DEFAULT_CONFIG_PATH}.
  --profile <name>         Database profile name in the local config. Defaults to the config default profile.
  --publish-command <cmd>  Publish command. Defaults to "${DEFAULT_PUBLISH_COMMAND}".
  --branch <name>          Release branch override. Defaults to v{major}.{minor}.x.
  --extract-cwd <path>     Optional source scan cwd override. Omit to scan the target package directory.
  --origin-path <path>     Per-package origin snapshot path. Defaults to <package-dir>/locales/origin.txt.
  --output-dir <path>      Agent-managed Excel output directory. Defaults to ${path.join(DEFAULT_AGENT_HOME, 'outputs')}.
  --modules <modules>      Legacy module selector for unusual repo-root scans. Omit for normal package releases.
  --reserved-keys <keys>   Reserved key prefixes for source scan.
  --lang <lang>            Source language. Defaults to en_US.
  --translate-lang <lang>  Target language queried from DB. Defaults to zh_CN.
  --commit-message <msg>   Commit message.
  --commit-file <path>     Extra file to include in commit. Can be repeated.

DB init options:
  --init-db-config         Write a local DB config template. Does not connect to the database.
  --force                  Overwrite an existing DB config when used with --init-db-config.
  --db-type <type>         Currently supports mysql.
  --db-host <host>
  --db-port <port>
  --db-user <username>
  --db-password <password> DB password stored in the local config file.
  --db-password-env <env>  Optional environment variable name that stores the DB password.
  --db-extra-where <sql>   Extra SQL condition in addition to tenant filtering.
  --db-batch-size <size>   Changed key query batch size. Defaults to ${DEFAULT_DB_BATCH_SIZE}.

Other:
  --json                   Print only JSON summary.
  --help                   Show this help.
`);
}

function parseArgs(argv) {
  const args = {
    commitFiles: [],
  };

  for (let index = 0; index < argv.length; index += 1) {
    const arg = argv[index];
    const next = () => {
      index += 1;
      if (index >= argv.length) {
        throw new Error(`Missing value for ${arg}`);
      }
      return argv[index];
    };

    switch (arg) {
      case '--target':
        args.target = next();
        break;
      case '--project':
        args.project = next();
        break;
      case '--package':
        args.package = next();
        break;
      case '--repo-root':
        args.repoRoot = next();
        break;
      case '--package-dir':
        args.packageDir = next();
        break;
      case '--version':
        args.version = next();
        break;
      case '--config':
        args.configPath = next();
        break;
      case '--profile':
        args.profile = next();
        break;
      case '--publish-command':
        args.publishCommand = next();
        break;
      case '--branch':
        args.branch = next();
        break;
      case '--extract-cwd':
        args.extractCwd = next();
        break;
      case '--origin-path':
        args.originPath = next();
        break;
      case '--output-dir':
        args.outputDir = next();
        break;
      case '--modules':
        args.modules = next();
        break;
      case '--reserved-keys':
        args.reservedKeys = next();
        break;
      case '--lang':
        args.lang = next();
        break;
      case '--translate-lang':
        args.translateLang = next();
        break;
      case '--commit-message':
        args.commitMessage = next();
        break;
      case '--commit-file':
        args.commitFiles.push(next());
        break;
      case '--execute':
        args.execute = true;
        break;
      case '--allow-publish':
        args.allowPublish = true;
        break;
      case '--allow-branch-switch':
        args.allowBranchSwitch = true;
        break;
      case '--allow-pull':
        args.allowPull = true;
        break;
      case '--allow-commit':
        args.allowCommit = true;
        break;
      case '--allow-push':
        args.allowPush = true;
        break;
      case '--skip-branch':
        args.skipBranch = true;
        break;
      case '--skip-pull':
        args.skipPull = true;
        break;
      case '--skip-publish':
        args.skipPublish = true;
        break;
      case '--skip-extract':
        args.skipExtract = true;
        break;
      case '--skip-commit':
        args.skipCommit = true;
        break;
      case '--skip-push':
        args.skipPush = true;
        break;
      case '--json':
        args.json = true;
        break;
      case '--init-db-config':
        args.initDbConfig = true;
        break;
      case '--force':
        args.force = true;
        break;
      case '--db-type':
        args.dbType = next();
        break;
      case '--db-host':
        args.dbHost = next();
        break;
      case '--db-port':
        args.dbPort = next();
        break;
      case '--db-name':
        args.dbName = next();
        break;
      case '--db-user':
        args.dbUser = next();
        break;
      case '--db-password':
        args.dbPassword = next();
        break;
      case '--db-password-env':
        args.dbPasswordEnv = next();
        break;
      case '--prompt-table':
        args.promptTable = next();
        break;
      case '--prompt-key-column':
        args.promptKeyColumn = next();
        break;
      case '--prompt-code-column':
        args.promptCodeColumn = next();
        break;
      case '--prompt-lang-column':
        args.promptLangColumn = next();
        break;
      case '--prompt-description-column':
        args.promptDescriptionColumn = next();
        break;
      case '--prompt-tenant-column':
        args.promptTenantColumn = next();
        break;
      case '--db-extra-where':
        args.dbExtraWhere = next();
        break;
      case '--db-batch-size':
        args.dbBatchSize = Number(next());
        break;
      case '--help':
      case '-h':
        args.help = true;
        break;
      default:
        if (!arg.startsWith('-') && !args.target && arg.includes('@')) {
          args.target = arg;
          break;
        }
        throw new Error(`Unknown argument: ${arg}`);
    }
  }

  return applyTargetShortcut(args);
}

function parseReleaseTarget(target) {
  const value = String(target || '').trim();
  const separatorIndex = value.lastIndexOf('@');
  if (separatorIndex <= 0 || separatorIndex === value.length - 1) {
    throw new Error(`Invalid release target: ${target}. Use package@version, for example package-name@1.9.1-beta.0.`);
  }
  return {
    package: value.slice(0, separatorIndex),
    version: value.slice(separatorIndex + 1),
  };
}

function applyTargetShortcut(args) {
  if (!args.target) {
    return args;
  }
  const parsed = parseReleaseTarget(args.target);
  if (args.package && args.package !== parsed.package) {
    throw new Error(`--target package (${parsed.package}) conflicts with --package (${args.package}).`);
  }
  if (args.version && args.version !== parsed.version) {
    throw new Error(`--target version (${parsed.version}) conflicts with --version (${args.version}).`);
  }
  return {
    ...args,
    package: args.package || parsed.package,
    version: args.version || parsed.version,
  };
}

function toAbsolutePath(inputPath, basePath = process.cwd()) {
  if (!inputPath) {
    return '';
  }
  const normalizedInput = String(inputPath).startsWith('~/')
    ? path.join(os.homedir(), String(inputPath).slice(2))
    : inputPath;
  return path.isAbsolute(normalizedInput) ? path.normalize(normalizedInput) : path.resolve(basePath, normalizedInput);
}

function readJson(filePath) {
  return JSON.parse(fs.readFileSync(filePath, 'utf8'));
}

function readJsonIfExists(filePath) {
  return fs.existsSync(filePath) ? readJson(filePath) : {};
}

function writeJson(filePath, data) {
  fs.mkdirSync(path.dirname(filePath), { recursive: true });
  fs.writeFileSync(filePath, `${JSON.stringify(data, null, 2)}\n`, 'utf8');
}

function writeText(filePath, text) {
  fs.mkdirSync(path.dirname(filePath), { recursive: true });
  fs.writeFileSync(filePath, text, 'utf8');
}

function firstDefined(...values) {
  return values.find((value) => value !== undefined && value !== null && value !== '');
}

function normalizeStringList(value) {
  if (value === undefined || value === null || value === '') {
    return undefined;
  }
  if (Array.isArray(value)) {
    return value.map((item) => String(item).trim()).filter(Boolean);
  }
  return String(value).split(',').map((item) => item.trim()).filter(Boolean);
}

function stringifyListOption(value) {
  const list = normalizeStringList(value);
  return list && list.length > 0 ? list.join(',') : undefined;
}

function sanitizePathSegment(value, fallback = 'default') {
  return String(value || fallback).replace(/[^A-Za-z0-9._-]+/g, '_');
}

function parseVersionParts(version) {
  const match = /^(\d+)\.(\d+)\.(\d+)(?:[-+].*)?$/.exec(version || '');
  if (!match) {
    return undefined;
  }
  return {
    major: match[1],
    minor: match[2],
    patch: match[3],
    majorMinor: `${match[1]}.${match[2]}`,
    majorMinorPatch: `${match[1]}.${match[2]}.${match[3]}`,
  };
}

function renderTemplate(template, context) {
  if (!template) {
    return template;
  }
  return String(template).replace(/\{([A-Za-z0-9_]+)\}/g, (source, key) => {
    const value = context[key];
    return value === undefined || value === null ? source : String(value);
  });
}

function buildVersionContext(version, extra = {}) {
  return {
    ...(parseVersionParts(version) || {}),
    version,
    ...extra,
  };
}

function resolveBranchOverride(version, configs, context) {
  const candidates = [
    version,
    context.majorMinorPatch,
    `${context.majorMinor}.x`,
    context.majorMinor,
  ].filter(Boolean);

  for (const config of configs) {
    const overrides = config?.branchOverrides || config?.branches;
    if (overrides && typeof overrides === 'object' && !Array.isArray(overrides)) {
      const matchedKey = candidates.find((candidate) => overrides[candidate]);
      if (matchedKey) {
        return {
          branch: renderTemplate(overrides[matchedKey], context),
          source: `branchOverrides.${matchedKey}`,
        };
      }
    }

    const rules = Array.isArray(config?.branchRules) ? config.branchRules : [];
    for (const rule of rules) {
      const ruleVersion = rule.version || rule.matchVersion;
      const ruleBranch = rule.branch || rule.pattern;
      if (!ruleBranch) {
        continue;
      }
      if (ruleVersion && candidates.includes(ruleVersion)) {
        return {
          branch: renderTemplate(ruleBranch, context),
          source: 'branchRules.version',
        };
      }
      if (rule.match && new RegExp(rule.match).test(version)) {
        return {
          branch: renderTemplate(ruleBranch, context),
          source: 'branchRules.match',
        };
      }
    }
  }

  return undefined;
}

function resolveReleaseBranch({ version, args, packageConfig, projectConfig, rootConfig, projectKey, packageKey }) {
  const versionParts = parseVersionParts(version);
  if (!versionParts) {
    return { branch: args.branch, source: args.branch ? 'argument' : '' };
  }

  const context = buildVersionContext(version, {
    projectKey: projectKey || '',
    packageKey: packageKey || '',
  });
  if (args.branch) {
    return { branch: renderTemplate(args.branch, context), source: 'argument' };
  }

  const configs = [packageConfig, projectConfig, rootConfig?.defaults].filter(Boolean);
  const override = resolveBranchOverride(version, configs, context);
  if (override) {
    return override;
  }

  const pattern = firstDefined(
    packageConfig?.releaseBranchPattern,
    packageConfig?.branchPattern,
    projectConfig?.releaseBranchPattern,
    projectConfig?.branchPattern,
    rootConfig?.defaults?.releaseBranchPattern,
    rootConfig?.defaults?.branchPattern,
    DEFAULT_RELEASE_BRANCH_PATTERN,
  );
  return {
    branch: renderTemplate(pattern, context),
    source: pattern === DEFAULT_RELEASE_BRANCH_PATTERN ? 'defaultPattern' : 'configuredPattern',
  };
}

function resolveVersionedWorkspaceRepoRoot({ version, packageDir, packageConfig, projectConfig, projectKey, packageKey }) {
  const workspaceConfig = packageConfig?.versionedWorkspace
    || packageConfig?.versionedWorkspaces
    || projectConfig?.versionedWorkspace
    || projectConfig?.versionedWorkspaces;
  if (!workspaceConfig) {
    return undefined;
  }

  const config = Array.isArray(workspaceConfig) ? workspaceConfig[0] : workspaceConfig;
  const baseDir = config.baseDir || config.rootDir || config.workspaceRoot;
  if (!baseDir) {
    throw new Error(`versionedWorkspace.baseDir is required for project ${projectKey || '(manual)'}`);
  }

  const context = buildVersionContext(version, {
    projectKey: projectKey || '',
    packageKey: packageKey || '',
  });
  const dirPattern = config.dirPattern || config.workspacePattern || DEFAULT_RELEASE_BRANCH_PATTERN;
  const repoPath = config.repoPath || config.projectPath || '.';
  const workspaceDir = toAbsolutePath(renderTemplate(dirPattern, context), toAbsolutePath(baseDir));
  const repoRoot = toAbsolutePath(renderTemplate(repoPath, context), workspaceDir);
  const packagePath = packageDir ? toAbsolutePath(packageDir, repoRoot) : '';

  if (!fs.existsSync(repoRoot) || !fs.statSync(repoRoot).isDirectory()) {
    throw new Error(`Versioned workspace repo does not exist: ${repoRoot}`);
  }
  if (!fs.existsSync(path.join(repoRoot, '.git'))) {
    throw new Error(`Versioned workspace is not a git repository: ${repoRoot}`);
  }
  if (packagePath && (!fs.existsSync(packagePath) || !fs.statSync(packagePath).isDirectory())) {
    throw new Error(`Package directory does not exist in versioned workspace: ${packagePath}`);
  }

  return {
    repoRoot,
    workspaceDir,
    source: 'versionedWorkspace',
  };
}

function findProjectByPackage(projects, packageKey) {
  const matches = Object.entries(projects)
    .filter(([, projectConfig]) => projectConfig?.packages?.[packageKey])
    .map(([projectKey]) => projectKey);
  if (matches.length === 1) {
    return matches[0];
  }
  if (matches.length > 1) {
    throw new Error(`Package key ${packageKey} exists in multiple projects: ${matches.join(', ')}. Use --project.`);
  }
  return undefined;
}

function resolveConfiguredTarget(args, rawConfig) {
  const projects = rawConfig.projects || rawConfig.workspaces || {};
  let projectKey = args.project;
  let packageKey = args.package;

  if (!projectKey && packageKey) {
    projectKey = findProjectByPackage(projects, packageKey);
  }
  if (!projectKey) {
    return {
      projectKey: '',
      packageKey: packageKey || '',
      projectConfig: {},
      packageConfig: {},
      source: 'manual',
    };
  }

  const projectConfig = projects[projectKey];
  if (!projectConfig) {
    throw new Error(`Project config not found: ${projectKey}`);
  }

  const packages = projectConfig.packages || {};
  if (!packageKey) {
    packageKey = projectConfig.defaultPackage;
  }

  if (!packageKey && Object.keys(packages).length === 1) {
    packageKey = Object.keys(packages)[0];
  }

  const packageConfig = packageKey ? packages[packageKey] : projectConfig.package;
  if (!packageConfig && packageKey) {
    throw new Error(`Package config not found: ${projectKey}/${packageKey}`);
  }
  if (!packageConfig && !projectConfig.packageDir) {
    throw new Error(`Package is required for project ${projectKey}. Use --package or set defaultPackage.`);
  }

  return {
    projectKey,
    packageKey: packageKey || '',
    projectConfig,
    packageConfig: packageConfig || {},
    source: 'config',
  };
}

function resolveSeedDataConfig({ packageConfig, projectConfig, rootConfig }) {
  const rootSeedData = rootConfig?.defaults?.seedData || {};
  const projectSeedData = projectConfig?.seedData || {};
  const packageSeedData = packageConfig?.seedData || {};
  const enabled = firstDefined(packageSeedData.enabled, projectSeedData.enabled, rootSeedData.enabled, true);
  const workbook = firstDefined(
    packageSeedData.workbook,
    packageSeedData.workbookPath,
    packageSeedData.path,
    projectSeedData.workbook,
    projectSeedData.workbookPath,
    projectSeedData.path,
    rootSeedData.workbook,
    rootSeedData.workbookPath,
    rootSeedData.path,
  );

  if (enabled === false || !workbook) {
    return undefined;
  }

  return {
    enabled: true,
    repoRoot: firstDefined(packageSeedData.repoRoot, projectSeedData.repoRoot, rootSeedData.repoRoot),
    workbook,
    worksheet: firstDefined(
      packageSeedData.worksheet,
      packageSeedData.sheetName,
      projectSeedData.worksheet,
      projectSeedData.sheetName,
      rootSeedData.worksheet,
      rootSeedData.sheetName,
      DEFAULT_SEED_WORKSHEET_NAME,
    ),
  };
}

function resolveReleaseArgs(args) {
  const configPath = toAbsolutePath(args.configPath || DEFAULT_CONFIG_PATH);
  const rawConfig = readJsonIfExists(configPath);
  const target = resolveConfiguredTarget(args, rawConfig);
  const projectExtract = target.projectConfig.extract || {};
  const packageExtract = target.packageConfig.extract || {};
  const releaseBranch = resolveReleaseBranch({
    version: args.version,
    args,
    packageConfig: target.packageConfig,
    projectConfig: target.projectConfig,
    rootConfig: rawConfig,
    projectKey: target.projectKey,
    packageKey: target.packageKey,
  });
  const packageDir = firstDefined(args.packageDir, target.packageConfig.packageDir, target.projectConfig.packageDir);
  const explicitRepoRoot = firstDefined(args.repoRoot, target.packageConfig.repoRoot);
  const fallbackRepoRoot = target.projectConfig.repoRoot;
  const versionedWorkspace = explicitRepoRoot ? undefined : resolveVersionedWorkspaceRepoRoot({
    version: args.version,
    packageDir,
    packageConfig: target.packageConfig,
    projectConfig: target.projectConfig,
    projectKey: target.projectKey,
    packageKey: target.packageKey,
  });
  const repoRoot = explicitRepoRoot || versionedWorkspace?.repoRoot || fallbackRepoRoot;
  const extractCwd = firstDefined(args.extractCwd, packageExtract.cwd, target.packageConfig.extractCwd);
  const extractModules = firstDefined(args.modules, packageExtract.modules, target.packageConfig.modules);

  return {
    ...args,
    configPath,
    project: target.projectKey || args.project,
    package: target.packageKey || args.package,
    configTargetSource: target.source,
    repoRoot,
    repoRootSource: explicitRepoRoot ? 'explicit' : (versionedWorkspace?.source || (fallbackRepoRoot ? 'configured' : '')),
    versionedWorkspace,
    packageDir,
    publishCommand: firstDefined(args.publishCommand, target.packageConfig.publishCommand, target.projectConfig.publishCommand, rawConfig.defaults?.publishCommand),
    extractCwd,
    originPath: firstDefined(args.originPath, packageExtract.originPath, target.packageConfig.originPath, projectExtract.originPath, target.projectConfig.originPath, rawConfig.defaults?.originPath),
    outputDir: firstDefined(args.outputDir, packageExtract.outputDir, target.packageConfig.outputDir, projectExtract.outputDir, target.projectConfig.outputDir, rawConfig.defaults?.outputDir),
    modules: stringifyListOption(extractModules),
    reservedKeys: stringifyListOption(firstDefined(args.reservedKeys, packageExtract.reservedKeys, target.packageConfig.reservedKeys, projectExtract.reservedKeys, target.projectConfig.reservedKeys)),
    lang: firstDefined(args.lang, packageExtract.lang, projectExtract.lang, rawConfig.defaults?.lang),
    translateLang: firstDefined(args.translateLang, packageExtract.translateLang, projectExtract.translateLang, rawConfig.defaults?.translateLang),
    releaseBranch: releaseBranch.branch,
    releaseBranchSource: releaseBranch.source,
    seedData: resolveSeedDataConfig({
      packageConfig: target.packageConfig,
      projectConfig: target.projectConfig,
      rootConfig: rawConfig,
    }),
  };
}

function assertFixedArchitectureArgs(args) {
  const mismatches = [];
  if (args.dbName && args.dbName !== FIXED_PROMPT_SCHEMA.database) {
    mismatches.push(`--db-name must be ${FIXED_PROMPT_SCHEMA.database}`);
  }
  if (args.promptTable && args.promptTable !== FIXED_PROMPT_SCHEMA.table) {
    mismatches.push(`--prompt-table must be ${FIXED_PROMPT_SCHEMA.table}`);
  }
  [
    ['--prompt-key-column', args.promptKeyColumn, FIXED_PROMPT_SCHEMA.columns.promptKey],
    ['--prompt-code-column', args.promptCodeColumn, FIXED_PROMPT_SCHEMA.columns.promptCode],
    ['--prompt-lang-column', args.promptLangColumn, FIXED_PROMPT_SCHEMA.columns.lang],
    ['--prompt-description-column', args.promptDescriptionColumn, FIXED_PROMPT_SCHEMA.columns.description],
    ['--prompt-tenant-column', args.promptTenantColumn, FIXED_PROMPT_SCHEMA.columns.tenantId],
  ].forEach(([flag, actual, expected]) => {
    if (actual && actual !== expected) {
      mismatches.push(`${flag} must be ${expected}`);
    }
  });
  if (mismatches.length > 0) {
    throw new Error(`Prompt table is fixed by company architecture: ${mismatches.join('; ')}`);
  }
}

function initializeDbConfig(args) {
  const configPath = toAbsolutePath(args.configPath || DEFAULT_CONFIG_PATH);
  if (fs.existsSync(configPath) && !args.force) {
    throw new Error(`DB config already exists: ${configPath}. Add --force to overwrite.`);
  }
  assertFixedArchitectureArgs(args);

  const profileName = args.profile || DEFAULT_PROFILE_NAME;
  const existingConfig = readJsonIfExists(configPath);
  const config = {
    ...existingConfig,
    defaultProfile: profileName,
    fixedPromptSchema: FIXED_PROMPT_SCHEMA,
    profiles: {
      ...(existingConfig.profiles || {}),
      [profileName]: {
        database: {
          type: args.dbType || 'mysql',
          host: args.dbHost || '<host>',
          port: Number(args.dbPort || 3306),
          username: args.dbUser || '<username>',
          password: args.dbPassword || '<password>',
          passwordEnv: args.dbPasswordEnv || '',
        },
        prompt: {
          extraWhere: args.dbExtraWhere || '',
        },
        batchSize: args.dbBatchSize || DEFAULT_DB_BATCH_SIZE,
      },
    },
  };

  writeJson(configPath, config);
  return configPath;
}

function assertDirectory(directoryPath, label) {
  if (!fs.existsSync(directoryPath) || !fs.statSync(directoryPath).isDirectory()) {
    throw new Error(`${label} does not exist or is not a directory: ${directoryPath}`);
  }
}

function assertFile(filePath, label) {
  if (!fs.existsSync(filePath) || !fs.statSync(filePath).isFile()) {
    throw new Error(`${label} does not exist or is not a file: ${filePath}`);
  }
}

function isPlaceholder(value) {
  return typeof value === 'string' && /^<.+>$/.test(value);
}

function pickConfigProfile(rawConfig, profileName) {
  if (!rawConfig.profiles) {
    return {
      profile: profileName || DEFAULT_PROFILE_NAME,
      database: rawConfig.database || {},
      prompt: rawConfig.prompt || {},
      batchSize: rawConfig.batchSize,
    };
  }

  const selectedProfile = profileName || rawConfig.defaultProfile || DEFAULT_PROFILE_NAME;
  const profile = rawConfig.profiles[selectedProfile];
  if (!profile) {
    throw new Error(`DB config profile not found: ${selectedProfile}`);
  }
  return {
    profile: selectedProfile,
    database: profile.database || {},
    prompt: profile.prompt || {},
    batchSize: profile.batchSize ?? rawConfig.batchSize,
  };
}

function normalizeDbConfig(rawConfig, configPath, profileName) {
  const selected = pickConfigProfile(rawConfig, profileName);
  return {
    configPath,
    profile: selected.profile,
    database: {
      type: selected.database.type || 'mysql',
      host: selected.database.host,
      port: selected.database.port || 3306,
      database: FIXED_PROMPT_SCHEMA.database,
      username: selected.database.username,
      password: selected.database.password,
      passwordEnv: selected.database.passwordEnv || '',
    },
    prompt: {
      table: FIXED_PROMPT_SCHEMA.table,
      columns: FIXED_PROMPT_SCHEMA.columns,
      tenantId: FIXED_TENANT_ID,
      extraWhere: selected.prompt.extraWhere || '',
    },
    batchSize: selected.batchSize || DEFAULT_DB_BATCH_SIZE,
  };
}

function loadDbConfig(configPath, profileName) {
  const resolvedPath = toAbsolutePath(configPath || DEFAULT_CONFIG_PATH);
  assertFile(resolvedPath, 'DB config');
  const config = normalizeDbConfig(readJson(resolvedPath), resolvedPath, profileName);
  validateDbConfig(config);
  return config;
}

function validateDbConfig(config) {
  const database = config.database || {};
  const prompt = config.prompt || {};
  const columns = prompt.columns || {};
  const required = [
    ['database.type', database.type],
    ['database.host', database.host],
    ['database.port', database.port],
    ['database.database', database.database],
    ['database.username', database.username],
    ['prompt.table', prompt.table],
    ['prompt.columns.promptKey', columns.promptKey],
    ['prompt.columns.promptCode', columns.promptCode],
    ['prompt.columns.lang', columns.lang],
    ['prompt.columns.description', columns.description],
    ['prompt.columns.tenantId', columns.tenantId],
  ];

  const missing = required
    .filter(([, value]) => value === undefined || value === null || value === '' || isPlaceholder(value))
    .map(([name]) => name);
  if (missing.length > 0) {
    throw new Error(`DB config is incomplete: ${missing.join(', ')}. Run --init-db-config with real values or edit the config file.`);
  }
  const hasConfigPassword = database.password !== undefined && database.password !== null && database.password !== '' && !isPlaceholder(database.password);
  const hasPasswordEnv = database.passwordEnv !== undefined && database.passwordEnv !== null && database.passwordEnv !== '' && !isPlaceholder(database.passwordEnv);
  if (!hasConfigPassword && !hasPasswordEnv) {
    throw new Error('DB config is incomplete: database.password or database.passwordEnv is required.');
  }
  if (database.type !== 'mysql') {
    throw new Error(`Unsupported database type: ${database.type}. Currently only mysql is supported.`);
  }
}

function detectPackageManager(repoRoot) {
  if (fs.existsSync(path.join(repoRoot, 'pnpm-lock.yaml'))) {
    return 'pnpm';
  }
  if (fs.existsSync(path.join(repoRoot, 'yarn.lock'))) {
    return 'yarn';
  }
  return 'npm';
}

function getCurrentBranch(repoRoot) {
  const result = runCommandCapture('git branch --show-current', repoRoot);
  return result.status === 0 ? result.stdout : '';
}

function buildBranchPlan(repoRoot, args) {
  const targetBranch = args.releaseBranch;
  const currentBranch = getCurrentBranch(repoRoot);
  const enabled = Boolean(targetBranch && !args.skipBranch && targetBranch !== currentBranch);
  const pullEnabled = Boolean(!args.skipPull && targetBranch && (!args.skipBranch || targetBranch === currentBranch));
  return {
    enabled,
    allowed: Boolean(args.allowBranchSwitch),
    cwd: repoRoot,
    current: currentBranch,
    target: targetBranch || '',
    source: args.releaseBranchSource || '',
    pull: {
      enabled: pullEnabled,
      allowed: Boolean(args.allowPull),
      command: 'git pull --ff-only',
      cwd: repoRoot,
    },
  };
}

function findExtractTarget({ repoRoot, packageDir, extractCwd }) {
  if (extractCwd) {
    const cwd = toAbsolutePath(extractCwd, repoRoot);
    return { cwd, source: 'explicit' };
  }

  if (fs.existsSync(path.join(packageDir, 'src'))) {
    return { cwd: packageDir, source: 'package' };
  }

  return { cwd: repoRoot, source: 'repo-root' };
}

function buildSeedDataPlan(seedData, repoRoot) {
  if (!seedData?.enabled || !seedData.workbook) {
    return undefined;
  }

  const seedRepoRoot = seedData.repoRoot ? toAbsolutePath(seedData.repoRoot, repoRoot) : repoRoot;
  const workbookPath = toAbsolutePath(seedData.workbook, seedRepoRoot);
  assertDirectory(seedRepoRoot, 'Seed data repository root');
  assertFile(workbookPath, 'Seed data workbook');

  return {
    enabled: true,
    repoRoot: seedRepoRoot,
    workbookPath,
    workbook: isInsideDirectory(workbookPath, seedRepoRoot) ? path.relative(seedRepoRoot, workbookPath) : workbookPath,
    worksheet: seedData.worksheet || DEFAULT_SEED_WORKSHEET_NAME,
    mode: 'upsert',
    uniqueKey: [
      SEED_COLUMN_KEYS.tenantId,
      SEED_COLUMN_KEYS.promptKey,
      SEED_COLUMN_KEYS.promptCode,
      SEED_COLUMN_KEYS.lang,
    ],
    writeColumns: [
      SEED_COLUMN_KEYS.promptKey,
      SEED_COLUMN_KEYS.promptCode,
      SEED_COLUMN_KEYS.lang,
      SEED_COLUMN_KEYS.description,
    ],
    autoCommit: false,
    timing: 'after-diff-before-origin-update',
  };
}

function runCommand(command, cwd) {
  const result = spawnSync(command, {
    cwd,
    shell: true,
    stdio: 'inherit',
    env: process.env,
  });

  if (result.status !== 0) {
    throw new Error(`Command failed (${result.status}): ${command}`);
  }
}

function runPublishCommand(command, cwd) {
  try {
    runCommand(command, cwd);
  } catch (error) {
    const whoami = runCommandCapture('npm whoami', cwd);
    const loginMessage = whoami.status === 0
      ? `Current npm login user: ${whoami.stdout || '(empty)'}`
      : `Unable to read npm login user after publish failure: ${whoami.stderr || whoami.stdout || 'npm whoami failed'}`;
    throw new Error(`${error.message}\n${loginMessage}`);
  }
}

function runCommandCapture(command, cwd) {
  const result = spawnSync(command, {
    cwd,
    shell: true,
    encoding: 'utf8',
    stdio: ['ignore', 'pipe', 'pipe'],
    env: process.env,
  });
  return {
    status: result.status,
    stdout: result.stdout.trim(),
    stderr: result.stderr.trim(),
  };
}

function normalizeCommitFile(filePath, repoRoot) {
  const absolutePath = toAbsolutePath(filePath, repoRoot);
  return path.relative(repoRoot, absolutePath);
}

function isInsideDirectory(filePath, directoryPath) {
  const relativePath = path.relative(directoryPath, toAbsolutePath(filePath, directoryPath));
  return Boolean(relativePath) && !relativePath.startsWith('..') && !path.isAbsolute(relativePath);
}

function quotePathspecList(filePaths) {
  return filePaths.map((filePath) => JSON.stringify(filePath)).join(' ');
}

function parseCsv(value) {
  return normalizeStringList(value);
}

function parseExtractOptions(args) {
  const modules = parseCsv(args.modules);
  if (modules?.includes('all')) {
    throw new Error('--modules all is no longer supported for package-scoped releases. Omit --modules so the scanner uses the target package directory.');
  }
  return {
    modules,
    reservedKeys: parseCsv(args.reservedKeys),
    lang: args.lang || 'en_US',
    translateLang: args.translateLang || 'zh_CN',
  };
}

function parseOriginText(text = '') {
  const map = new Map();
  text.split(/\r?\n/).forEach((line) => {
    if (!line.trim()) {
      return;
    }
    const separatorIndex = line.indexOf(':');
    if (separatorIndex < 0) {
      return;
    }
    const key = line.slice(0, separatorIndex);
    const value = line.slice(separatorIndex + 1);
    map.set(key, value);
  });
  return map;
}

function buildOriginText(localeList) {
  return localeList
    .slice()
    .sort((left, right) => left.key.localeCompare(right.key))
    .map((item) => [item.key, String(item.content || '').split('\n').join('')].join(':'))
    .join('\n');
}

function splitPromptKey(fullKey) {
  const [key1 = '', key2 = '', ...others] = String(fullKey).split('.');
  return {
    promptKey: [key1, key2].filter(Boolean).join('.'),
    promptCode: others.join('.'),
  };
}

function shouldSkipSourceEntry(entryName) {
  return entryName.startsWith('.')
    || ['node_modules', 'dist', 'build', 'doc'].includes(entryName)
    || /\.test\.[^.]+$/.test(entryName)
    || /\.config\.[^.]+$/.test(entryName);
}

function collectSourceFiles(directoryPath) {
  if (!fs.existsSync(directoryPath) || !fs.statSync(directoryPath).isDirectory()) {
    return [];
  }

  const files = [];
  fs.readdirSync(directoryPath).forEach((entryName) => {
    if (shouldSkipSourceEntry(entryName)) {
      return;
    }
    const entryPath = path.join(directoryPath, entryName);
    const stat = fs.lstatSync(entryPath);
    if (stat.isDirectory()) {
      files.push(...collectSourceFiles(entryPath));
    } else if (SOURCE_EXTENSIONS.includes(path.extname(entryPath))) {
      files.push(entryPath);
    }
  });
  return files;
}

function transformDefaultMessage(message, shouldTrim) {
  if (!shouldTrim) {
    return message;
  }
  return String(message).replace(/\${([a-zA-Z0-9_\s\t]+)}/gm, '{$1}');
}

function extractMessagesFromContent(content, filePath) {
  const messages = [];
  let match;
  const detectRegexp = new RegExp(DETECT_REGEXP);
  while ((match = detectRegexp.exec(content)) != null) {
    const keyIndex = match[3] ? 3 : 6;
    const messageIndex = match[5] ? 5 : 8;
    const key = match[keyIndex];
    const defaultMessage = match[messageIndex];
    const shouldTrim = /\.(d|defaultMessage)\([\s\S]*`[\s\S]*\)/.test(match[0]);
    messages.push({
      key,
      path: filePath,
      originalDefaultMessage: defaultMessage,
      transformedDefaultMessage: transformDefaultMessage(defaultMessage, shouldTrim),
    });
  }

  const noDefaultRegexp = new RegExp(NO_DEFAULT_REGEXP);
  while ((match = noDefaultRegexp.exec(content)) != null) {
    messages.push({
      key: match[2],
      path: filePath,
      isValid: false,
      invalidType: 'no_default',
    });
  }
  return messages;
}

function scanSourceMessages(sourcePath) {
  return collectSourceFiles(sourcePath).flatMap((filePath) => {
    const content = fs.readFileSync(filePath, 'utf8');
    return extractMessagesFromContent(content, filePath);
  });
}

function verifySourceMessages(messages) {
  const errors = [];
  const seen = new Map();
  messages.forEach((message) => {
    if (message.isValid === false && message.invalidType === 'no_default') {
      errors.push(`The key="${message.key}" has no default message ${message.path}`);
      return;
    }
    if (!message.key || message.originalDefaultMessage == null) {
      return;
    }
    const previous = seen.get(message.key);
    if (previous && previous.originalDefaultMessage !== message.originalDefaultMessage) {
      errors.push(`The key="${message.key}" has different default message "${message.originalDefaultMessage}" ${message.path}`);
    } else {
      seen.set(message.key, message);
    }
  });

  if (errors.length > 0) {
    throw new Error(`Source i18n scan failed:\n${errors.join('\n')}`);
  }
}

function resolvePackageModuleConfigs(cwd, modules) {
  const singlePackageJsonPath = path.join(cwd, 'package.json');
  const singlePackageSourcePath = path.join(cwd, 'src');
  const hasSinglePackageLayout = fs.existsSync(singlePackageJsonPath) && fs.existsSync(singlePackageSourcePath);
  const buildSinglePackageConfig = () => {
    const pkg = readJson(singlePackageJsonPath);
    return {
      name: pkg.name,
      version: pkg.version,
      moduleName: pkg.name,
      sourcePath: singlePackageSourcePath,
    };
  };
  const buildPkgPath = (moduleName) => path.join(cwd, 'packages', moduleName, 'package.json');
  const buildModuleConfig = (moduleName) => {
    const pkg = readJson(buildPkgPath(moduleName));
    return {
      moduleName,
      name: pkg.name,
      version: pkg.version,
      sourcePath: path.join(cwd, 'packages', moduleName, 'src'),
    };
  };

  if (Array.isArray(modules) && modules.length > 0) {
    return modules
      .filter((moduleName) => fs.existsSync(buildPkgPath(moduleName)))
      .map(buildModuleConfig);
  }

  const pkg = readJson(path.join(cwd, 'package.json'));
  return [{
    name: pkg.name,
    version: pkg.version,
    moduleName: pkg.name,
    sourcePath: path.join(cwd, 'src'),
  }];
}

function runCurrentSourceScan(plan) {
  const moduleConfigs = resolvePackageModuleConfigs(plan.commands.extract.cwd, plan.extractOptions.modules);
  const reservedKeys = plan.extractOptions.reservedKeys || [];
  const sourceObj = {};

  moduleConfigs.forEach((moduleConfig) => {
    const messages = scanSourceMessages(moduleConfig.sourcePath)
      .filter((item) => {
        if (reservedKeys.length === 0) {
          return true;
        }
        const prefix = item.key.split('.')[0];
        return reservedKeys.some((key) => prefix === key);
      });
    verifySourceMessages(messages);
    messages
      .forEach((item) => {
        sourceObj[item.key] = {
          name: moduleConfig.name,
          version: moduleConfig.version,
          moduleName: moduleConfig.moduleName,
          key: item.key,
          content: String(item.transformedDefaultMessage || '').split('\n').join(''),
        };
      });
  });

  return Object.keys(sourceObj).sort().map((key) => sourceObj[key]);
}

function quoteMysqlIdentifier(identifier) {
  return `\`${String(identifier).replace(/`/g, '``')}\``;
}

function quoteMysqlQualifiedName(name) {
  return String(name).split('.').map(quoteMysqlIdentifier).join('.');
}

function quoteMysqlValue(value) {
  return `'${String(value).replace(/\\/g, '\\\\').replace(/'/g, "''")}'`;
}

function chunkArray(list, size) {
  const chunks = [];
  for (let index = 0; index < list.length; index += size) {
    chunks.push(list.slice(index, index + size));
  }
  return chunks;
}

function buildPromptWhere(items, columns) {
  const grouped = new Map();
  items.forEach((item) => {
    const { promptKey, promptCode } = splitPromptKey(item.key);
    if (!promptKey) {
      return;
    }
    if (!grouped.has(promptKey)) {
      grouped.set(promptKey, new Set());
    }
    grouped.get(promptKey).add(promptCode);
  });

  const promptKeyColumn = quoteMysqlIdentifier(columns.promptKey);
  const promptCodeColumn = quoteMysqlIdentifier(columns.promptCode);
  const clauses = Array.from(grouped.entries()).map(([promptKey, promptCodes]) => {
    const codeList = Array.from(promptCodes).map(quoteMysqlValue).join(', ');
    return `(${promptKeyColumn} = ${quoteMysqlValue(promptKey)} AND ${promptCodeColumn} IN (${codeList}))`;
  });

  return clauses.length > 0 ? `(${clauses.join(' OR ')})` : '(1 = 0)';
}

function buildPromptQuery({ config, items, languages }) {
  const columns = config.prompt.columns;
  const selectedColumns = [
    `${quoteMysqlIdentifier(columns.promptKey)} AS prompt_key`,
    `${quoteMysqlIdentifier(columns.promptCode)} AS prompt_code`,
    `${quoteMysqlIdentifier(columns.lang)} AS lang`,
    `${quoteMysqlIdentifier(columns.description)} AS description`,
  ].join(', ');

  const where = [
    buildPromptWhere(items, columns),
    `${quoteMysqlIdentifier(columns.lang)} IN (${languages.map(quoteMysqlValue).join(', ')})`,
  ];
  if (columns.tenantId && config.prompt.tenantId !== undefined && config.prompt.tenantId !== null && config.prompt.tenantId !== '') {
    where.push(`${quoteMysqlIdentifier(columns.tenantId)} = ${quoteMysqlValue(config.prompt.tenantId)}`);
  }
  if (config.prompt.extraWhere) {
    where.push(`(${config.prompt.extraWhere})`);
  }

  return `SELECT ${selectedColumns} FROM ${quoteMysqlQualifiedName(config.prompt.table)} WHERE ${where.join(' AND ')}`;
}

function queryMysqlPrompts({ config, items, languages }) {
  if (items.length === 0) {
    return new Map();
  }

  const database = config.database;
  const batchSize = Number(config.batchSize || DEFAULT_DB_BATCH_SIZE);
  const envPassword = database.passwordEnv ? process.env[database.passwordEnv] : undefined;
  const configPassword = isPlaceholder(database.password) ? '' : database.password;
  const password = envPassword || configPassword || '';
  const promptMap = new Map();

  if (!password) {
    throw new Error('DB password is not configured. Set database.password in the local config or provide database.passwordEnv.');
  }

  chunkArray(items, batchSize).forEach((batch) => {
    const sql = buildPromptQuery({ config, items: batch, languages });
    const mysqlArgs = [
      '--no-defaults',
      `--host=${database.host}`,
      `--port=${database.port}`,
      `--user=${database.username}`,
      `--database=${database.database}`,
      '--batch',
      '--raw',
      '--skip-column-names',
      '--default-character-set=utf8mb4',
      `--execute=${sql}`,
    ];
    const result = spawnSync('mysql', mysqlArgs, {
      shell: false,
      encoding: 'utf8',
      stdio: ['ignore', 'pipe', 'pipe'],
      env: {
        ...process.env,
        MYSQL_PWD: password,
      },
    });

    if (result.error) {
      throw new Error(`Failed to run mysql client: ${result.error.message}. Install mysql client or make it available in PATH.`);
    }
    if (result.status !== 0) {
      throw new Error(`MySQL query failed (${result.status}): ${result.stderr || result.stdout}`);
    }

    result.stdout.split(/\r?\n/).filter(Boolean).forEach((line) => {
      const [promptKey, promptCode, lang, description = ''] = line.split('\t');
      const fullKey = [promptKey, promptCode].filter(Boolean).join('.');
      if (!promptMap.has(fullKey)) {
        promptMap.set(fullKey, {});
      }
      promptMap.get(fullKey)[lang] = description;
    });
  });

  return promptMap;
}

function loadSkillDependency(dependencyName) {
  try {
    return SKILL_REQUIRE(dependencyName);
  } catch (error) {
    throw new Error(
      `Skill runtime dependency "${dependencyName}" is unavailable. Run "npm install" in ${SKILL_ROOT}. ${error.message}`,
    );
  }
}

async function writePromptExcel({ outputPath, diffItems, promptMap, lang, translateLang }) {
  const ExcelJS = loadSkillDependency('exceljs');
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('\u591a\u8bed\u8a00\u5bfc\u51fa');
  worksheet.columns = [
    { header: '\u6a21\u677f\u4ee3\u7801', key: 'prompt_key', width: 20 },
    { header: '\u4ee3\u7801', key: 'prompt_code', width: 30 },
    { header: '\u63cf\u8ff0(\u4e2d\u6587)', key: 'lang_zh', width: 50 },
    { header: '\u63cf\u8ff0(English)', key: 'lang_en', width: 50 },
  ];

  const rows = diffItems.map((item) => {
    const { promptKey, promptCode } = splitPromptKey(item.key);
    const dbText = promptMap.get(item.key) || {};
    const zh = dbText.zh_CN || '';
    const en = dbText.en_US || '';
    return [promptKey, promptCode, zh, en];
  });

  worksheet.addRows(rows);
  fs.mkdirSync(path.dirname(outputPath), { recursive: true });
  await workbook.xlsx.writeFile(outputPath);
}

function getExcelCellText(cell) {
  const value = cell?.value;
  if (value === undefined || value === null) {
    return '';
  }
  if (value instanceof Date) {
    return value.toISOString();
  }
  if (typeof value === 'object') {
    if (Array.isArray(value.richText)) {
      return value.richText.map((item) => item.text || '').join('').trim();
    }
    if (value.text !== undefined && value.text !== null) {
      return String(value.text).trim();
    }
    if (value.result !== undefined && value.result !== null) {
      return String(value.result).trim();
    }
    if (value.formula) {
      return String(value.result || '').trim();
    }
  }
  return String(value).trim();
}

function normalizeSeedHeaderName(value) {
  return String(value || '').trim().replace(/^[#*]+/, '').toLowerCase();
}

function mapSeedHeaderKey(value) {
  const normalized = normalizeSeedHeaderName(value);
  return Object.entries(SEED_COLUMN_KEYS).find(([, columnName]) => columnName === normalized)?.[0];
}

function tryFindSeedHeader(worksheet) {
  const maxHeaderRow = Math.min(worksheet.rowCount, 50);
  for (let rowNumber = 1; rowNumber <= maxHeaderRow; rowNumber += 1) {
    const row = worksheet.getRow(rowNumber);
    const columns = {};
    row.eachCell({ includeEmpty: false }, (cell, columnNumber) => {
      const key = mapSeedHeaderKey(getExcelCellText(cell));
      if (key) {
        columns[key] = columnNumber;
      }
    });

    const requiredKeys = [
      'tenantId',
      'promptKey',
      'promptCode',
      'lang',
      'description',
    ];
    if (requiredKeys.every((key) => columns[key])) {
      return {
        rowNumber,
        columns,
      };
    }
  }
  return undefined;
}

function locateSeedWorksheet(workbook, worksheetName) {
  const configuredWorksheet = worksheetName ? workbook.getWorksheet(worksheetName) : undefined;
  const candidates = [
    configuredWorksheet,
    ...workbook.worksheets.filter((worksheet) => worksheet !== configuredWorksheet),
  ].filter(Boolean);

  for (const worksheet of candidates) {
    const header = tryFindSeedHeader(worksheet);
    if (header) {
      return {
        worksheet,
        header,
      };
    }
  }

  throw new Error(`Seed workbook worksheet not found or header is invalid. Expected sheet "${worksheetName || DEFAULT_SEED_WORKSHEET_NAME}" with prompt seed columns.`);
}

function buildSeedRowKey({ tenantId, promptKey, promptCode, lang }) {
  return [tenantId, promptKey, promptCode, lang].map((item) => String(item ?? '')).join('\u0001');
}

function buildSeedPromptEntries({ diffItems, promptMap, lang, translateLang }) {
  const languages = Array.from(new Set([lang, translateLang].filter(Boolean)));
  const entries = [];
  let skipped = 0;
  diffItems.forEach((item) => {
    const { promptKey, promptCode } = splitPromptKey(item.key);
    if (!promptKey) {
      return;
    }
    const dbText = promptMap.get(item.key) || {};
    languages.forEach((language) => {
      const description = dbText[language];
      if (description === undefined || description === null || String(description) === '') {
        skipped += 1;
        return;
      }
      entries.push({
        tenantId: String(FIXED_TENANT_ID),
        promptKey,
        promptCode,
        lang: language,
        description: String(description),
      });
    });
  });
  return {
    entries,
    skipped,
  };
}

function escapeRegExp(value) {
  return String(value).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function decodeXmlAttribute(value) {
  return String(value || '')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&amp;/g, '&');
}

function escapeXmlText(value) {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

function escapeXmlAttribute(value) {
  return escapeXmlText(value)
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function parseXmlAttributes(tag) {
  const attributes = {};
  const attributePattern = /([A-Za-z_][\w:.-]*)="([^"]*)"/g;
  let match;
  while ((match = attributePattern.exec(tag))) {
    attributes[match[1]] = decodeXmlAttribute(match[2]);
  }
  return attributes;
}

function formatXmlAttributes(attributes) {
  return Object.entries(attributes)
    .filter(([, value]) => value !== undefined && value !== null)
    .map(([key, value]) => `${key}="${escapeXmlAttribute(value)}"`)
    .join(' ');
}

function columnNumberToName(columnNumber) {
  let value = Number(columnNumber);
  let name = '';
  while (value > 0) {
    const remainder = (value - 1) % 26;
    name = String.fromCharCode(65 + remainder) + name;
    value = Math.floor((value - 1) / 26);
  }
  return name;
}

function columnNameToNumber(columnName) {
  return String(columnName || '').toUpperCase().split('').reduce((sum, char) => {
    const code = char.charCodeAt(0);
    if (code < 65 || code > 90) {
      return sum;
    }
    return sum * 26 + code - 64;
  }, 0);
}

function splitCellRef(ref) {
  const match = String(ref || '').match(/^([A-Za-z]+)(\d+)$/);
  if (!match) {
    return undefined;
  }
  return {
    columnName: match[1].toUpperCase(),
    columnNumber: columnNameToNumber(match[1]),
    rowNumber: Number(match[2]),
  };
}

function buildCellRef(columnNumber, rowNumber) {
  return `${columnNumberToName(columnNumber)}${rowNumber}`;
}

function buildInlineStringXml(value) {
  const text = String(value ?? '');
  const preserveSpace = /^\s|\s$|[\r\n\t]/.test(text) ? ' xml:space="preserve"' : '';
  return `<is><t${preserveSpace}>${escapeXmlText(text)}</t></is>`;
}

function buildInlineStringCellXml(ref, value, sourceAttributes = {}) {
  const attributes = {
    ...sourceAttributes,
    r: ref,
    t: 'inlineStr',
  };
  return `<c ${formatXmlAttributes(attributes)}>${buildInlineStringXml(value)}</c>`;
}

function replaceCellValueXml(cellXml, value) {
  const openTag = cellXml.match(/^<c\b[^>]*(?:\/>|>)/)?.[0] || '<c>';
  const attributes = parseXmlAttributes(openTag);
  return buildInlineStringCellXml(attributes.r, value, attributes);
}

function getRowXml(sheetXml, rowNumber) {
  const rowPattern = new RegExp(`<row\\b(?=[^>]*\\br="${rowNumber}"(?:\\D|"))[^>]*>[\\s\\S]*?<\\/row>`);
  return sheetXml.match(rowPattern)?.[0];
}

function getLastWorksheetRowNumber(sheetXml) {
  let lastRowNumber = 0;
  const rowPattern = /<row\b[^>]*\br="(\d+)"[^>]*>/g;
  let match;
  while ((match = rowPattern.exec(sheetXml))) {
    lastRowNumber = Math.max(lastRowNumber, Number(match[1]));
  }
  return lastRowNumber;
}

function collectCellAttributesByColumn(rowXml) {
  const attributesByColumn = new Map();
  if (!rowXml) {
    return attributesByColumn;
  }
  const cellPattern = /<c\b[^>]*(?:\/>|>[\s\S]*?<\/c>)/g;
  let match;
  while ((match = cellPattern.exec(rowXml))) {
    const openTag = match[0].match(/^<c\b[^>]*(?:\/>|>)/)?.[0];
    const attributes = parseXmlAttributes(openTag || '');
    const ref = splitCellRef(attributes.r);
    if (!ref) {
      continue;
    }
    delete attributes.r;
    delete attributes.t;
    attributesByColumn.set(ref.columnNumber, attributes);
  }
  return attributesByColumn;
}

function collectRowTemplateAttributes(rowXml) {
  if (!rowXml) {
    return {};
  }
  const openTag = rowXml.match(/^<row\b[^>]*>/)?.[0];
  const attributes = parseXmlAttributes(openTag || '');
  delete attributes.r;
  return attributes;
}

function insertCellXmlIntoRow(rowXml, columnNumber, cellXml) {
  const cellPattern = /<c\b[^>]*(?:\/>|>[\s\S]*?<\/c>)/g;
  let match;
  while ((match = cellPattern.exec(rowXml))) {
    const openTag = match[0].match(/^<c\b[^>]*(?:\/>|>)/)?.[0];
    const attributes = parseXmlAttributes(openTag || '');
    const ref = splitCellRef(attributes.r);
    if (ref && ref.columnNumber > columnNumber) {
      return `${rowXml.slice(0, match.index)}${cellXml}${rowXml.slice(match.index)}`;
    }
  }
  return rowXml.replace('</row>', `${cellXml}</row>`);
}

function patchRowCellValue(rowXml, rowNumber, columnNumber, value, templateCellAttributesByColumn) {
  const ref = buildCellRef(columnNumber, rowNumber);
  const cellPattern = new RegExp(`<c\\b(?=[^>]*\\br="${escapeRegExp(ref)}"(?=\\D|"))[^>]*(?:\\/>|>[\\s\\S]*?<\\/c>)`);
  const existingCellXml = rowXml.match(cellPattern)?.[0];
  if (existingCellXml) {
    return rowXml.replace(cellPattern, replaceCellValueXml(existingCellXml, value));
  }

  const sourceAttributes = templateCellAttributesByColumn.get(columnNumber) || {};
  const cellXml = buildInlineStringCellXml(ref, value, sourceAttributes);
  return insertCellXmlIntoRow(rowXml, columnNumber, cellXml);
}

function buildSeedDataRowXml({ rowNumber, valuesByColumn, templateRowAttributes, templateCellAttributesByColumn }) {
  const rowAttributes = {
    ...templateRowAttributes,
    r: String(rowNumber),
  };
  const cells = Array.from(valuesByColumn.entries())
    .sort(([leftColumn], [rightColumn]) => leftColumn - rightColumn)
    .map(([columnNumber, value]) => buildInlineStringCellXml(
      buildCellRef(columnNumber, rowNumber),
      value,
      templateCellAttributesByColumn.get(columnNumber) || {},
    ));
  return `<row ${formatXmlAttributes(rowAttributes)}>${cells.join('')}</row>`;
}

function updateWorksheetDimension(sheetXml, maxRowNumber, maxColumnNumber) {
  return sheetXml.replace(/<dimension\b[^>]*\bref="([^"]+)"[^>]*\/>/, (dimensionXml, ref) => {
    const [startRef, endRef = startRef] = ref.split(':');
    const end = splitCellRef(endRef);
    const nextMaxRow = Math.max(end?.rowNumber || 0, maxRowNumber);
    const nextMaxColumn = Math.max(end?.columnNumber || 0, maxColumnNumber);
    return dimensionXml.replace(ref, `${startRef}:${buildCellRef(nextMaxColumn, nextMaxRow)}`);
  });
}

function updateWorksheetAutoFilter(sheetXml, maxRowNumber, maxColumnNumber) {
  return sheetXml.replace(/<autoFilter\b[^>]*\bref="([^"]+)"[^>]*(?:\/>|>[\s\S]*?<\/autoFilter>)/, (autoFilterXml, ref) => {
    const [startRef, endRef = startRef] = ref.split(':');
    const end = splitCellRef(endRef);
    const nextMaxRow = Math.max(end?.rowNumber || 0, maxRowNumber);
    const nextMaxColumn = Math.max(end?.columnNumber || 0, maxColumnNumber);
    return autoFilterXml.replace(ref, `${startRef}:${buildCellRef(nextMaxColumn, nextMaxRow)}`);
  });
}

async function resolveWorksheetXmlPathFromZip(zip, worksheetName) {
  const workbookXml = await zip.file('xl/workbook.xml')?.async('string');
  const workbookRelsXml = await zip.file('xl/_rels/workbook.xml.rels')?.async('string');
  if (!workbookXml || !workbookRelsXml) {
    throw new Error('Seed workbook is missing xl/workbook.xml or workbook relationships.');
  }

  let relationshipId = '';
  const sheetPattern = /<sheet\b[^>]*\/>/g;
  let sheetMatch;
  while ((sheetMatch = sheetPattern.exec(workbookXml))) {
    const attributes = parseXmlAttributes(sheetMatch[0]);
    if (attributes.name === worksheetName) {
      relationshipId = attributes['r:id'];
      break;
    }
  }
  if (!relationshipId) {
    throw new Error(`Seed workbook worksheet XML not found for "${worksheetName}".`);
  }

  const relPattern = /<Relationship\b[^>]*\/>/g;
  let relMatch;
  while ((relMatch = relPattern.exec(workbookRelsXml))) {
    const attributes = parseXmlAttributes(relMatch[0]);
    if (attributes.Id !== relationshipId) {
      continue;
    }
    const target = attributes.Target || '';
    return target.startsWith('/')
      ? target.replace(/^\/+/, '')
      : path.posix.normalize(path.posix.join('xl', target));
  }

  throw new Error(`Seed workbook worksheet relationship not found for "${worksheetName}".`);
}

async function patchSeedPromptWorkbookXml({ workbookPath, worksheetName, updates, inserts, header }) {
  const JSZip = loadSkillDependency('jszip');
  const zip = await JSZip.loadAsync(fs.readFileSync(workbookPath));
  const worksheetXmlPath = await resolveWorksheetXmlPathFromZip(zip, worksheetName);
  const worksheetFile = zip.file(worksheetXmlPath);
  if (!worksheetFile) {
    throw new Error(`Seed workbook worksheet XML file not found: ${worksheetXmlPath}`);
  }

  let sheetXml = await worksheetFile.async('string');
  const firstDataRowXml = getRowXml(sheetXml, header.rowNumber + 1);
  const lastExistingRowNumber = getLastWorksheetRowNumber(sheetXml);
  const lastDataRowXml = getRowXml(sheetXml, lastExistingRowNumber) || firstDataRowXml;
  const templateRowAttributes = collectRowTemplateAttributes(lastDataRowXml);
  const templateCellAttributesByColumn = collectCellAttributesByColumn(lastDataRowXml);
  const firstDataCellAttributesByColumn = collectCellAttributesByColumn(firstDataRowXml);
  firstDataCellAttributesByColumn.forEach((attributes, columnNumber) => {
    if (!templateCellAttributesByColumn.has(columnNumber)) {
      templateCellAttributesByColumn.set(columnNumber, attributes);
    }
  });

  const updatesByRow = new Map();
  updates.forEach(({ rowNumber, columnNumber, value }) => {
    if (!updatesByRow.has(rowNumber)) {
      updatesByRow.set(rowNumber, []);
    }
    updatesByRow.get(rowNumber).push({ columnNumber, value });
  });

  updatesByRow.forEach((rowUpdates, rowNumber) => {
    const originalRowXml = getRowXml(sheetXml, rowNumber);
    if (!originalRowXml) {
      throw new Error(`Seed workbook row ${rowNumber} not found while updating existing prompt.`);
    }
    const patchedRowXml = rowUpdates.reduce(
      (rowXml, { columnNumber, value }) => patchRowCellValue(rowXml, rowNumber, columnNumber, value, templateCellAttributesByColumn),
      originalRowXml,
    );
    sheetXml = sheetXml.replace(originalRowXml, patchedRowXml);
  });

  if (inserts.length > 0) {
    const insertRowsXml = inserts.map(({ rowNumber, valuesByColumn }) => buildSeedDataRowXml({
      rowNumber,
      valuesByColumn,
      templateRowAttributes,
      templateCellAttributesByColumn,
    })).join('');
    sheetXml = sheetXml.replace('</sheetData>', `${insertRowsXml}</sheetData>`);
  }

  const maxRowNumber = Math.max(
    lastExistingRowNumber,
    ...updates.map((item) => item.rowNumber),
    ...inserts.map((item) => item.rowNumber),
  );
  const maxColumnNumber = Math.max(
    0,
    ...Object.values(header.columns).filter(Boolean),
  );
  sheetXml = updateWorksheetDimension(sheetXml, maxRowNumber, maxColumnNumber);
  sheetXml = updateWorksheetAutoFilter(sheetXml, maxRowNumber, maxColumnNumber);

  zip.file(worksheetXmlPath, sheetXml);
  const output = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
  fs.writeFileSync(workbookPath, output);
}

async function writeSeedPromptWorkbook({ seedData, diffItems, promptMap, lang, translateLang }) {
  const summary = {
    enabled: Boolean(seedData?.enabled),
    workbookPath: seedData?.workbookPath || '',
    worksheet: seedData?.worksheet || DEFAULT_SEED_WORKSHEET_NAME,
    updated: 0,
    inserted: 0,
    unchanged: 0,
    skipped: 0,
    changedRows: 0,
  };

  if (!seedData?.enabled || diffItems.length === 0) {
    return summary;
  }

  const ExcelJS = loadSkillDependency('exceljs');
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(seedData.workbookPath);
  const { worksheet, header } = locateSeedWorksheet(workbook, seedData.worksheet);
  summary.worksheet = worksheet.name;

  const existingRows = new Map();
  for (let rowNumber = header.rowNumber + 1; rowNumber <= worksheet.rowCount; rowNumber += 1) {
    const row = worksheet.getRow(rowNumber);
    const promptKey = getExcelCellText(row.getCell(header.columns.promptKey));
    const promptCode = getExcelCellText(row.getCell(header.columns.promptCode));
    const rowLang = getExcelCellText(row.getCell(header.columns.lang));
    if (!promptKey || !rowLang) {
      continue;
    }
    existingRows.set(buildSeedRowKey({
      tenantId: getExcelCellText(row.getCell(header.columns.tenantId)) || String(FIXED_TENANT_ID),
      promptKey,
      promptCode,
      lang: rowLang,
    }), {
      row,
      rowNumber,
    });
  }

  let nextRowNumber = worksheet.rowCount + 1;
  const updates = [];
  const inserts = [];
  const seedEntries = buildSeedPromptEntries({ diffItems, promptMap, lang, translateLang });
  summary.skipped += seedEntries.skipped;
  seedEntries.entries.forEach((entry) => {
    const rowKey = buildSeedRowKey(entry);
    const existingRow = existingRows.get(rowKey);
    if (existingRow) {
      const { row, rowNumber } = existingRow;
      const currentDescription = getExcelCellText(row.getCell(header.columns.description));
      if (currentDescription === entry.description) {
        summary.unchanged += 1;
        return;
      }
      updates.push({
        rowNumber,
        columnNumber: header.columns.description,
        value: entry.description,
      });
      summary.updated += 1;
      return;
    }

    const valuesByColumn = new Map();
    if (header.columns.promptId) {
      valuesByColumn.set(header.columns.promptId, '*');
    }
    valuesByColumn.set(header.columns.tenantId, entry.tenantId);
    valuesByColumn.set(header.columns.promptKey, entry.promptKey);
    valuesByColumn.set(header.columns.promptCode, entry.promptCode);
    valuesByColumn.set(header.columns.lang, entry.lang);
    valuesByColumn.set(header.columns.description, entry.description);
    if (header.columns.lastUpdateDate) {
      valuesByColumn.set(header.columns.lastUpdateDate, SEED_WORKBOOK_LAST_UPDATE_VALUE);
    }
    if (header.columns.lastUpdatedBy) {
      valuesByColumn.set(header.columns.lastUpdatedBy, SEED_WORKBOOK_LAST_UPDATED_BY);
    }
    inserts.push({
      rowNumber: nextRowNumber,
      valuesByColumn,
    });
    existingRows.set(rowKey, {
      row: undefined,
      rowNumber: nextRowNumber,
    });
    nextRowNumber += 1;
    summary.inserted += 1;
  });

  summary.changedRows = summary.updated + summary.inserted;
  if (summary.changedRows > 0) {
    await patchSeedPromptWorkbookXml({
      workbookPath: seedData.workbookPath,
      worksheetName: worksheet.name,
      updates,
      inserts,
      header,
    });
  }
  return summary;
}

function buildI18nStatePaths({ params, repoRoot, packageDir, packageName, branchName }) {
  const outputRoot = toAbsolutePath(params.outputDir || path.join(DEFAULT_AGENT_HOME, 'outputs'));
  const projectSegment = sanitizePathSegment(params.project || path.basename(repoRoot), 'project');
  const packageSegment = sanitizePathSegment(params.package || packageName || path.basename(repoRoot), 'package');
  const branchSegment = sanitizePathSegment(branchName || params.version, 'branch');
  const relativeDir = path.join(projectSegment, packageSegment, branchSegment);
  const originPath = params.originPath
    ? toAbsolutePath(renderTemplate(params.originPath, {
      version: params.version,
      projectKey: params.project || '',
      packageKey: params.package || '',
      packageName: packageName || '',
      packageDir: path.relative(repoRoot, packageDir),
    }), repoRoot)
    : path.join(packageDir, 'locales', 'origin.txt');
  const legacyOriginPath = path.join(repoRoot, 'locales', 'origin.txt');
  const originSourcePath = fs.existsSync(originPath)
    ? originPath
    : (originPath !== legacyOriginPath && fs.existsSync(legacyOriginPath) ? legacyOriginPath : originPath);
  return {
    originPath,
    originSourcePath,
    originFallbackPath: originSourcePath !== originPath ? originSourcePath : '',
    excelPath: path.join(outputRoot, relativeDir, EXCEL_OUTPUTS[0]),
    outputRoot,
  };
}

async function runDatabaseExtraction(plan) {
  const originSourcePath = plan.i18nState.originSourcePath || plan.i18nState.originPath;
  const originMap = fs.existsSync(originSourcePath) ? parseOriginText(fs.readFileSync(originSourcePath, 'utf8')) : new Map();
  const currentList = runCurrentSourceScan(plan);
  const diffItems = currentList.filter((item) => originMap.get(item.key) !== item.content);
  const languages = Array.from(new Set([plan.extractOptions.lang, plan.extractOptions.translateLang].filter(Boolean)));
  const promptMap = queryMysqlPrompts({ config: plan.dbConfig, items: diffItems, languages });

  if (diffItems.length > 0) {
    await writePromptExcel({
      outputPath: plan.i18nState.excelPath,
      diffItems,
      promptMap,
      lang: plan.extractOptions.lang,
      translateLang: plan.extractOptions.translateLang,
    });
  }

  const seedSummary = await writeSeedPromptWorkbook({
    seedData: plan.seedData,
    diffItems,
    promptMap,
    lang: plan.extractOptions.lang,
    translateLang: plan.extractOptions.translateLang,
  });

  writeText(plan.i18nState.originPath, buildOriginText(currentList));
  return {
    outputs: [
      plan.i18nState.originPath,
      ...(diffItems.length > 0 ? [plan.i18nState.excelPath] : []),
      ...(seedSummary.changedRows > 0 ? [seedSummary.workbookPath] : []),
    ],
    diffSummary: {
      changed: diffItems.length,
      total: currentList.length,
      queriedLanguages: languages,
      origin: plan.i18nState.originPath,
      originSource: originSourcePath,
      diffBasis: 'origin.txt',
      sourceOfTruth: 'database-prompts',
      changeDetector: 'source-code-vs-origin',
      seedData: seedSummary,
    },
  };
}

function buildPlan(args) {
  const params = resolveReleaseArgs(args);
  if (!params.repoRoot) {
    throw new Error('--repo-root is required');
  }
  if (!params.packageDir) {
    throw new Error('--package-dir is required');
  }
  if (!params.version) {
    throw new Error('--version is required');
  }
  if (!/^\d+\.\d+\.\d+(?:[-+][0-9A-Za-z.-]+)?$/.test(params.version)) {
    throw new Error(`Invalid semantic version: ${params.version}`);
  }

  const repoRoot = toAbsolutePath(params.repoRoot);
  const packageDir = toAbsolutePath(params.packageDir, repoRoot);
  assertDirectory(repoRoot, 'Repository root');
  assertDirectory(packageDir, 'Package directory');

  const packageJsonPath = path.join(packageDir, 'package.json');
  assertFile(packageJsonPath, 'Package package.json');

  const rootPackageJsonPath = path.join(repoRoot, 'package.json');
  assertFile(rootPackageJsonPath, 'Root package.json');

  const packageJson = readJson(packageJsonPath);
  const packageManager = detectPackageManager(repoRoot);
  const extract = findExtractTarget({
    repoRoot,
    packageDir,
    extractCwd: params.extractCwd,
  });

  assertDirectory(extract.cwd, 'Extraction cwd');

  const dbConfig = params.skipExtract
    ? undefined
    : (fs.existsSync(params.configPath) ? loadDbConfig(params.configPath, params.profile) : undefined);
  if (!params.skipExtract && params.execute && !dbConfig) {
    throw new Error(`DB config does not exist: ${params.configPath}. Run --init-db-config first.`);
  }
  const extractOptions = parseExtractOptions(params);
  const templateContext = {
    version: params.version,
    projectKey: params.project || '',
    packageKey: params.package || '',
    packageName: packageJson.name || '',
    packageDir: path.relative(repoRoot, packageDir),
  };

  const publishCommand = renderTemplate(params.publishCommand || DEFAULT_PUBLISH_COMMAND, templateContext);
  const commitMessage = renderTemplate(
    params.commitMessage || `发布${packageJson.name || path.basename(packageDir)}@${params.version}`,
    templateContext,
  );
  const branchPlan = buildBranchPlan(repoRoot, params);
  const i18nState = buildI18nStatePaths({
    params,
    repoRoot,
    packageDir,
    packageName: packageJson.name,
    branchName: branchPlan.target,
  });
  const seedData = !params.skipExtract ? buildSeedDataPlan(params.seedData, repoRoot) : undefined;
  const defaultCommitFiles = [packageJsonPath];
  if (!params.skipExtract && isInsideDirectory(i18nState.originPath, repoRoot)) {
    defaultCommitFiles.push(i18nState.originPath);
  }
  const extraCommitFiles = params.commitFiles.map((filePath) => toAbsolutePath(filePath, repoRoot));
  const branchNeedsPreparation = branchPlan.enabled || branchPlan.pull.enabled;
  const currentVersion = branchNeedsPreparation ? '' : packageJson.version;
  const commandPlan = {
    branch: branchPlan,
    extract: {
      enabled: !params.skipExtract,
      mode: 'database',
      cwd: extract.cwd,
      source: extract.source,
      configPath: dbConfig?.configPath || params.configPath,
      profile: dbConfig?.profile || params.profile || '',
      options: extractOptions,
      timing: 'before-publish',
      continueWhenNoDiff: true,
      diffBasis: 'origin.txt',
      sourceOfTruth: 'database-prompts',
      changeDetector: 'source-code-vs-origin',
      seedData: seedData ? {
        enabled: true,
        repoRoot: seedData.repoRoot,
        workbook: seedData.workbookPath,
        worksheet: seedData.worksheet,
        mode: seedData.mode,
        uniqueKey: seedData.uniqueKey,
        writeColumns: seedData.writeColumns,
        autoCommit: seedData.autoCommit,
      } : undefined,
    },
    version: {
      enabled: true,
      file: normalizeCommitFile(packageJsonPath, repoRoot),
      current: currentVersion,
      currentSource: currentVersion ? 'current-worktree' : 'after-branch-prepared',
      target: params.version,
      timing: 'after-extract-before-publish',
    },
    publish: {
      enabled: !params.skipPublish,
      allowed: Boolean(params.allowPublish),
      command: publishCommand,
      cwd: packageDir,
    },
    commit: {
      enabled: !params.skipCommit,
      allowed: Boolean(params.allowCommit),
      message: commitMessage,
      files: [...defaultCommitFiles, ...extraCommitFiles].map((filePath) => normalizeCommitFile(filePath, repoRoot)),
    },
  };
  commandPlan.push = {
    enabled: commandPlan.commit.enabled && !params.skipPush,
    allowed: Boolean(params.allowPush),
    command: commandPlan.branch.target ? `git push origin HEAD:${commandPlan.branch.target}` : 'git push',
  };

  return {
    execute: Boolean(params.execute),
    configTargetSource: params.configTargetSource,
    project: params.project || '',
    package: params.package || '',
    repoRoot,
    repoRootSource: params.repoRootSource || '',
    versionedWorkspace: params.versionedWorkspace,
    packageDir,
    packageJsonPath,
    packageName: packageJson.name,
    currentVersion: commandPlan.version.current,
    currentVersionSource: commandPlan.version.currentSource,
    targetVersion: params.version,
    packageManager,
    dbConfig,
    extractOptions,
    seedData,
    commands: commandPlan,
    i18nState,
    expectedI18nOutputs: [
      i18nState.originPath,
      i18nState.excelPath,
      ...(seedData ? [seedData.workbookPath] : []),
    ],
  };
}

function validateExecution(plan) {
  if (!plan.execute) {
    return;
  }
  if (plan.commands.branch.enabled && !plan.commands.branch.allowed) {
    throw new Error('Branch switch is needed but not allowed. Add --allow-branch-switch or --skip-branch.');
  }
  if (plan.commands.branch.pull.enabled && !plan.commands.branch.pull.allowed) {
    throw new Error('Git pull is enabled but not allowed. Add --allow-pull or --skip-pull.');
  }
  if (plan.commands.publish.enabled && !plan.commands.publish.allowed) {
    throw new Error('Publish is enabled but not allowed. Add --allow-publish or --skip-publish.');
  }
  if (plan.commands.commit.enabled && !plan.commands.commit.allowed) {
    throw new Error('Commit is enabled but not allowed. Add --allow-commit or --skip-commit.');
  }
  if (plan.commands.push.enabled && !plan.commands.push.allowed) {
    throw new Error('Push is enabled but not allowed. Add --allow-push or --skip-push.');
  }
}

function printPlan(plan) {
  const currentVersion = plan.currentVersion || `read after ${plan.commands.branch.target || 'branch'} is prepared`;
  console.log(`Release i18n agent plan:
- package: ${plan.packageName || path.basename(plan.packageDir)}
- project/package key: ${[plan.project, plan.package].filter(Boolean).join('/') || 'manual'}
- repo root: ${plan.repoRoot}${plan.repoRootSource ? ` (${plan.repoRootSource})` : ''}
- package dir: ${plan.packageDir}
- version: ${currentVersion} -> ${plan.targetVersion}
- mode: ${plan.execute ? 'execute' : 'dry-run'}
- branch: ${plan.commands.branch.target ? `${plan.commands.branch.current || '(detached)'} -> ${plan.commands.branch.target} (${plan.commands.branch.source})` : 'not configured'}
- pull: ${plan.commands.branch.pull.enabled ? `${plan.commands.branch.pull.command} (cwd: ${plan.commands.branch.pull.cwd || plan.repoRoot})` : 'skipped'}
- extract: ${plan.commands.extract.enabled ? `built-in scanner (${plan.commands.extract.cwd})` : 'skipped'}
- diff basis: ${plan.commands.extract.enabled ? 'origin.txt only' : 'skipped'}
- change detector: ${plan.commands.extract.enabled ? 'source code vs origin.txt' : 'skipped'}
- final text source: ${plan.commands.extract.enabled ? 'hzero_platform.hpfm_prompt' : 'skipped'}
- origin: ${plan.commands.extract.enabled ? plan.i18nState.originPath : 'skipped'}
- origin source: ${plan.commands.extract.enabled ? (plan.i18nState.originFallbackPath || plan.i18nState.originPath) : 'skipped'}
- excel: ${plan.commands.extract.enabled ? `${plan.i18nState.excelPath} (generated only when diff exists)` : 'skipped'}
- seed workbook: ${plan.seedData ? `${plan.seedData.workbookPath} (upsert when diff exists; not auto-committed)` : 'skipped'}
- seed upsert key: ${plan.seedData ? plan.seedData.uniqueKey.join(', ') : 'skipped'}
- seed write columns: ${plan.seedData ? plan.seedData.writeColumns.join(', ') : 'skipped'}
- update version: ${plan.commands.version.file} ${currentVersion} -> ${plan.commands.version.target} (${plan.commands.version.timing})
- publish: ${plan.commands.publish.enabled ? plan.commands.publish.command : 'skipped'}
- commit: ${plan.commands.commit.enabled ? plan.commands.commit.message : 'skipped'}
- push: ${plan.commands.push.enabled ? plan.commands.push.command : 'skipped'}
`);
}

async function executePlan(plan) {
  if (plan.commands.branch.enabled) {
    runCommand(`git checkout ${JSON.stringify(plan.commands.branch.target)}`, plan.repoRoot);
  }
  if (plan.commands.branch.pull.enabled) {
    runCommand(plan.commands.branch.pull.command, plan.repoRoot);
  }

  let extractionResult = { outputs: [], diffSummary: undefined };
  if (plan.commands.extract.enabled) {
    extractionResult = await runDatabaseExtraction(plan);
    plan.commands.extract.completed = true;
    plan.commands.extract.changed = extractionResult.diffSummary?.changed ?? 0;
  }

  const packageJson = readJson(plan.packageJsonPath);
  plan.currentVersion = packageJson.version;
  plan.currentVersionSource = 'after-branch-prepared';
  plan.commands.version.current = packageJson.version;
  plan.commands.version.currentSource = 'after-branch-prepared';
  packageJson.version = plan.targetVersion;
  writeJson(plan.packageJsonPath, packageJson);
  plan.commands.version.updated = true;

  if (plan.commands.publish.enabled) {
    runPublishCommand(plan.commands.publish.command, plan.commands.publish.cwd);
  }

  if (plan.commands.commit.enabled) {
    const commitFiles = Array.from(new Set([
      ...plan.commands.commit.files,
      ...extractionResult.outputs
        .filter((filePath) => isInsideDirectory(filePath, plan.repoRoot))
        .map((filePath) => normalizeCommitFile(filePath, plan.repoRoot)),
    ]));
    runCommand(`git add -- ${quotePathspecList(commitFiles)}`, plan.repoRoot);
    runCommand(`git commit -m ${JSON.stringify(plan.commands.commit.message)} -- ${quotePathspecList(commitFiles)}`, plan.repoRoot);
    plan.commands.commit.files = commitFiles;
    const commitHash = runCommandCapture('git rev-parse HEAD', plan.repoRoot);
    if (commitHash.status === 0) {
      plan.commands.commit.hash = commitHash.stdout;
    }
  }

  if (plan.commands.push.enabled) {
    runCommand(plan.commands.push.command, plan.repoRoot);
    plan.commands.push.pushed = true;
  }

  return extractionResult;
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  if (args.help) {
    printHelp();
    return;
  }

  if (args.initDbConfig) {
    const configPath = initializeDbConfig(args);
    const initializedConfig = readJson(configPath);
    const profileName = initializedConfig.defaultProfile || DEFAULT_PROFILE_NAME;
    const summary = { initialized: true, configPath, profile: profileName };
    if (!args.json) {
      console.log(`DB config initialized: ${configPath}`);
      console.log('DB password is stored in the local config file. Keep this file private.');
    }
    console.log(JSON.stringify(summary, null, 2));
    return;
  }

  const plan = buildPlan(args);
  validateExecution(plan);

  let extractionResult = { outputs: [], diffSummary: undefined };
  if (!args.json) {
    printPlan(plan);
  }

  const gitStatusBefore = runCommandCapture('git status --short', plan.repoRoot);

  if (plan.execute) {
    extractionResult = await executePlan(plan);
  }

  const summary = {
    ...plan,
    dbConfig: plan.dbConfig ? {
      configPath: plan.dbConfig.configPath,
      profile: plan.dbConfig.profile,
      database: {
        type: plan.dbConfig.database.type,
        host: plan.dbConfig.database.host,
        port: plan.dbConfig.database.port,
        database: plan.dbConfig.database.database,
        username: plan.dbConfig.database.username,
        passwordEnv: plan.dbConfig.database.passwordEnv,
      },
      prompt: plan.dbConfig.prompt,
      batchSize: plan.dbConfig.batchSize,
    } : {
      configPath: plan.commands.extract.configPath,
      profile: plan.commands.extract.profile,
      missing: plan.commands.extract.enabled,
    },
    gitStatusBefore,
    i18nOutputs: plan.execute ? extractionResult.outputs : plan.expectedI18nOutputs,
    diffSummary: extractionResult.diffSummary,
  };

  console.log(JSON.stringify(summary, null, 2));
}

try {
  await main();
} catch (error) {
  console.error(error.message);
  process.exit(1);
}
