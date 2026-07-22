# Company Skills

This repository stores reusable Agent Skills for HZero-based product teams. Each skill is directly under `<repository-root>/<skill-name>/`; no application-specific adapter layer is required for the skill core.

## Structure

```text
<repository-root>/
└── hzero-release-i18n-agent/
    ├── SKILL.md
    ├── package.json
    ├── agents/openai.yaml
    ├── references/config.example.json
    └── scripts/release-i18n-agent.mjs
```

`agents/openai.yaml` is optional Codex UI metadata. Other Agent Skills-compatible tools can use `SKILL.md`, the bundled script, and the references directly.

## Available Skills

### hzero-release-i18n-agent

Release an HZero/JIPaaS frontend package from a chat request, export the database-backed multilingual diff before publishing, update the package-scoped `locales/origin.txt`, and upsert the configured backend seed-data workbook without committing that seed-data repository.

The release script is conservative: it defaults to dry-run and requires explicit execution flags for branch switching, pulling, publishing, committing, and pushing.

## Local Setup

Install the skill dependencies once:

```bash
cd /path/to/skills/hzero-release-i18n-agent
npm install
```

Create a private local config outside this repository. Copy and complete [`config.example.json`](hzero-release-i18n-agent/references/config.example.json), then restrict permissions:

```bash
mkdir -p "${XDG_CONFIG_HOME:-$HOME/.config}/hzero-release-i18n-agent"
cp /path/to/skills/hzero-release-i18n-agent/references/config.example.json \
  "${XDG_CONFIG_HOME:-$HOME/.config}/hzero-release-i18n-agent/config.json"
chmod 600 "${XDG_CONFIG_HOME:-$HOME/.config}/hzero-release-i18n-agent/config.json"
```

The config contains database credentials and team-specific workspace mappings. It is intentionally ignored by Git and must never be committed.

Each Agent product has its own skill-discovery setup. Point that product at `hzero-release-i18n-agent/` or install/copy this folder according to its Agent Skills mechanism; the skill workflow itself is not tied to Codex.

## Runtime Requirements

- Node.js 18 or later.
- `npm install` completed in the skill directory.
- Local `mysql` command available before database-backed extraction.
- A completed local configuration with HZero project/package/seed-data mappings.

No real release, database query, or seed-data repository change is performed merely by installing the skill.
