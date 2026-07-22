# Company Skills

This repository stores reusable Agent Skills for HZero-based product teams. Skills are listed under `skills/<skill-name>/` so standard skill installers can discover them.

## Structure

```text
<repository-root>/
└── skills/
    └── hzero-release-i18n-agent/
        ├── SKILL.md
        ├── package.json
        ├── agents/openai.yaml
        ├── references/config.example.json
        └── scripts/
            ├── release-i18n-agent.mjs
            └── setup.mjs
```

`agents/openai.yaml` is optional Codex UI metadata. The skill core remains `SKILL.md`, `scripts/`, and `references/`.

## Available Skills

### hzero-release-i18n-agent

Release an HZero/JIPaaS frontend package from a chat request, export the database-backed multilingual diff before publishing, update the package-scoped `locales/origin.txt`, and upsert the configured backend seed-data workbook without committing that seed-data repository.

The release script is conservative: it defaults to dry-run and requires explicit execution flags for branch switching, pulling, publishing, committing, and pushing.

## Install

Install the release skill with a standard skill installer:

```bash
npx skills@latest add Mafiti/skills --skill hzero-release-i18n-agent
```

The installer places the skill into the selected Agent's skill directory. It installs skill files only; it does not install the bundled Node runtime dependencies.

Before the first release dry-run, run the skill setup script from the installed skill directory:

```bash
SKILL_DIR=/path/to/installed/hzero-release-i18n-agent
node "$SKILL_DIR/scripts/setup.mjs"
```

`setup.mjs` is idempotent. It verifies `exceljs` and `jszip`, then runs `npm ci --omit=dev --ignore-scripts` only when they are missing.

Create a private local config outside the skill repository. Copy and complete `references/config.example.json` from the installed skill directory, then restrict permissions:

```bash
CONFIG_DIR="${XDG_CONFIG_HOME:-$HOME/.config}/hzero-release-i18n-agent"
mkdir -p "$CONFIG_DIR"
cp "$SKILL_DIR/references/config.example.json" "$CONFIG_DIR/config.json"
chmod 600 "$CONFIG_DIR/config.json"
```

The config contains database credentials and team-specific workspace mappings. It is intentionally ignored by Git and must never be committed.

## Runtime Requirements

- Node.js 18 or later.
- `node "$SKILL_DIR/scripts/setup.mjs"` completed in the skill directory.
- Local `mysql` command available before database-backed extraction.
- A completed local configuration with HZero project/package/seed-data mappings.

No real release, database query, or seed-data repository change is performed merely by installing the skill.
