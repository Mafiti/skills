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

## First Use

Start a new conversation with the installed Agent and say:

```text
初始化 hzero-release-i18n-agent
```

The Agent resolves its own installed skill directory, runs the idempotent runtime setup, creates the private config template, and then asks for the required database and repository mappings. Users do not need to find or set `SKILL_DIR`.

For manual terminal troubleshooting only, use the installation path printed by `npx skills` as `SKILL_DIR`:

```bash
SKILL_DIR=/path/printed/by/installer/hzero-release-i18n-agent
node "$SKILL_DIR/scripts/setup.mjs"
node "$SKILL_DIR/scripts/release-i18n-agent.mjs" --init-config
```

The private config is generated at `${XDG_CONFIG_HOME:-$HOME/.config}/hzero-release-i18n-agent/config.json`. It contains the database profile and `projects -> packages` mappings, stays outside Git, and must never be committed.

## Runtime Requirements

- Node.js 18 or later.
- Complete the Agent-led first-use initialization, or run the manual setup command above.
- Local `mysql` command available before database-backed extraction.
- A completed local configuration with HZero project/package/seed-data mappings.

No real release, database query, or seed-data repository change is performed merely by installing the skill.
