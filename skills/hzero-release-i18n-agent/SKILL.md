---
name: hzero-release-i18n-agent
description: Release HZero/JIPaaS frontend packages from chat requests, including shorthand such as "发布 package-name@1.2.3", release-branch resolution, package version updates, package-scoped multilingual origin snapshots, database-backed multilingual diff Excel export, seed-data workbook upsert, and frontend release commit/push. Use when an agent needs to initialize local release settings, prepare a release dry-run, publish a configured frontend package, export release multilingual differences, or synchronize prompt changes to backend seed data.
---

# HZero Release I18n Agent

Use this skill for a configured HZero/JIPaaS frontend package release. The skill can be invoked in chat or by running its bundled script; it is not tied to one Agent application.

## Required Workflow

1. Parse a request such as `发布 package-name@1.2.3` into package and target version.
2. Read the private local config and resolve the project, package directory, release branch, optional versioned workspace, source scan options, and seed workbook.
3. Inspect repository branch and dirty state. Do not trust the current `package.json` version when the configured release branch differs.
4. Run a dry-run and show the planned branch switch, pull, multilingual extraction, version update, publish, commit, and push.
5. Wait for explicit user confirmation before any side effect. A dry-run is never release authorization.
6. In execution mode, switch to the expected branch and pull from the repository root before extraction or publishing.
7. Extract multilingual differences before publishing. Use source code only as the current-key/change detector, query the database for final text, update the package-scoped `locales/origin.txt`, and upsert the configured seed workbook.
8. When extraction succeeds, update the package `version`, publish from the package directory, then commit and push only the frontend release repository files.
9. Report generated Excel path, origin snapshot path, seed workbook result, version change, publish result, commit hash, and push result.

## Safety Rules

- Default to dry-run. Never invent a target version.
- Require explicit confirmation in the current conversation before `git checkout`, `git pull`, package publishing, `git commit`, `git push`, or merge-request actions.
- Run all Git operations from the resolved repository root. Run the publish command only from the target package directory.
- Treat commit and push as part of confirmed release execution after successful publish. Do not add a second confirmation checkpoint unless the user asks.
- Do not run `npm whoami`, registry checks, or login commands before publishing. Assume local login already exists. Only inspect `npm whoami` after a publish failure.
- Multilingual extraction must finish before publishing. Zero changed keys is successful and must not block publishing.
- Update package `package.json` only after extraction succeeds and before publishing.
- Stage only release files in the frontend repository, normally the target package `package.json` and `<package-dir>/locales/origin.txt`. Never commit the generated diff Excel.
- Never automatically commit or push the seed-data repository. Workbook updates remain local changes for the user to review and commit separately.
- Treat local config as trusted code: repository paths, publish commands, and SQL conditions are executed locally. Do not use an unreviewed config.
- Do not pass database passwords on the command line. Store `database.password` or `database.passwordEnv` only in the private local config.

## Installation And Runtime Setup

This skill is intended to be discovered from a repository `skills/<skill-name>/` directory. A skill installer copies the skill files, but does not install the bundled Node dependencies.

Before the first release dry-run, set `SKILL_DIR` to the installed skill directory and run setup:

```bash
SKILL_DIR=/path/to/installed/hzero-release-i18n-agent
node "$SKILL_DIR/scripts/setup.mjs"
```

`setup.mjs` is idempotent. It checks for `exceljs` and `jszip`, then runs `npm ci --omit=dev --ignore-scripts` only when the dependencies are absent. Do not make the release script install dependencies implicitly.

The script defaults to:

- config: `${XDG_CONFIG_HOME:-$HOME/.config}/hzero-release-i18n-agent/config.json`
- output directory: `${XDG_STATE_HOME:-$HOME/.local/state}/hzero-release-i18n-agent/outputs`

Override either location with `--config` or `--output-dir`.

Create the private config by copying `references/config.example.json`. Complete database credentials and team-specific repository mappings, then restrict permissions:

```bash
CONFIG_DIR="${XDG_CONFIG_HOME:-$HOME/.config}/hzero-release-i18n-agent"
mkdir -p "$CONFIG_DIR"
cp "$SKILL_DIR/references/config.example.json" "$CONFIG_DIR/config.json"
chmod 600 "$CONFIG_DIR/config.json"
```

Alternatively, initialize a placeholder database profile and then edit the generated config:

```bash
node "$SKILL_DIR/scripts/release-i18n-agent.mjs" --init-db-config
```

The config is private runtime state. Never add it to a business repository or this skill repository.

## Commands

Configured dry-run:

```bash
node "$SKILL_DIR/scripts/release-i18n-agent.mjs" \
  --target package-name@1.2.3
```

Explicit dry-run:

```bash
node "$SKILL_DIR/scripts/release-i18n-agent.mjs" \
  --repo-root /path/to/frontend-repository \
  --package-dir packages/package-name \
  --version 1.2.3 \
  --reserved-keys promptPrefix \
  --config /private/path/config.json
```

Execute only after the user confirms the dry-run plan:

```bash
node "$SKILL_DIR/scripts/release-i18n-agent.mjs" \
  --target package-name@1.2.3 \
  --execute \
  --allow-branch-switch \
  --allow-pull \
  --allow-publish \
  --allow-commit \
  --allow-push
```

Use `--skip-publish` only when the user explicitly states that publication was already completed or intentionally skipped. The script still requires all applicable execution flags for branch, pull, commit, and push.

## Config Model

Use `projects -> packages` for reusable configuration:

- `projects.<project>.repoRoot`: direct local repository checkout.
- `projects.<project>.versionedWorkspace`: use when release branches have separate local checkouts. `dirPattern` defaults to `v{major}.{minor}.x`.
- `projects.<project>.branchPattern`: release branch pattern. The default maps `1.9.1-beta.0` to `v1.9.x`.
- `branchOverrides` or `branchRules`: configure only version-to-branch exceptions.
- `packages.<package>.packageDir`: package directory relative to the resolved repository root.
- `packages.<package>.extract.reservedKeys`: package-owned multilingual key prefixes.
- `seedData.repoRoot` and `seedData.workbook`: local backend seed-data workbook to update.

When a package key appears in exactly one configured project, `--target package@version` resolves that project automatically. If it appears in multiple projects, require `--project`.

See `references/config.example.json` for the complete portable shape. Keep real local paths and credentials only in the private copied config.

## Multilingual Rules

- Scan the target package source by default; do not call a project's legacy `intl:extract` or `locales/extract` implementation.
- Store the prior source snapshot at `<package-dir>/locales/origin.txt`.
- Compare current source keys/messages with that snapshot to identify changed keys. If source and snapshot conflict, update the snapshot from source; source is the change detector only.
- Query `hzero_platform.hpfm_prompt` for final descriptions. The database, not source default text, is the final text source for the exported Excel and seed workbook.
- Apply `reservedKeys` filtering before duplicate-key validation. Fail only when included package keys conflict.
- Generate the diff Excel outside the business repository only when differences exist. No difference means success.
- Upsert seed data by `tenant_id + prompt_key + prompt_code + lang`. Existing rows only update the business description; new rows inherit fixed template values. Never delete rows.
- Preserve seed workbook formatting. The script patches worksheet OOXML with `jszip`; do not replace that path with `workbook.xlsx.writeFile()`.

## Fixed HZero Database Model

This is an HZero-specific skill. The prompt query model is intentionally fixed:

- schema: `hzero_platform`
- table: `hpfm_prompt`
- key columns: `prompt_key`, `prompt_code`, `lang`, `tenant_id`
- text column: `description`
- tenant filter: `tenant_id = 0`

Ask only for database host, port, username, password or password environment variable, and profile name. Do not ask users to configure schema, table, column names, or tenant ID unless their product explicitly uses a non-standard architecture.

## Runtime Dependencies

The skill bundle owns `exceljs` and `jszip`; initialize them with `setup.mjs`. Database extraction additionally requires a local `mysql` command.

Before reporting readiness, run:

```bash
node --check "$SKILL_DIR/scripts/release-i18n-agent.mjs"
```

Before a real release, run at least one dry-run against the resolved repository. Do not claim that publishing, commit, push, database extraction, Excel export, or seed-data update succeeded unless it ran in the current turn.
