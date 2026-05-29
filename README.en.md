# historiographer-skills

An open repository of reusable AI skills for team collaboration, internal knowledge sharing, and community contribution.

Chinese version: [README.md](README.md)

## Purpose

This repository is designed to:

1. Turn personal skills into reusable assets
2. Provide a shared structure for team-maintained skills
3. Support open collaboration through the standard GitHub fork / pull request workflow

In this repository, a skill is usually a self-contained directory that includes at least one `SKILL.md`, and optionally scripts, templates, or references.

## Repository Structure

```text
historiographer-skills/
  ├── README.md
  ├── README.en.md
  ├── CONTRIBUTING.md
  ├── LICENSE
  └── skills/
      ├── docx-add-screenshots/
      │   ├── SKILL.md
      │   └── scripts/
      │       └── docx_screenshot_inserter.py
      └── wsl-win-openbrowser/
          ├── SKILL.md
          └── references/
              └── .gitkeep
```

## Skill Layout

Each skill should remain as self-contained as possible. A skill directory may include:

- `SKILL.md`
  The core description of the skill, including purpose, trigger conditions, workflow, constraints, and expected output
- `scripts/`
  Helper scripts required by the skill
- `references/`
  Notes, command examples, or supporting documentation
- `templates/`
  Reusable templates
- `assets/`
  Images or static resources

The minimum valid structure is:

```text
skills/my-skill/
  └── SKILL.md
```

## Included Skills

### `docx-add-screenshots`

Purpose:
Batch insert browser screenshots into specific sections of a Word document. Useful for operation manuals, product guides, and delivery documents.

Contents:

- `SKILL.md`
- `scripts/docx_screenshot_inserter.py`

### `wsl-win-openbrowser`

Purpose:
Drive browser operations from WSL through Windows-side `openCLI`, including opening pages, waiting for load, checking state, and taking screenshots.

Contents:

- `SKILL.md`
- `references/`

## How to Use

### 1. Use it as a reference repository

Browse the `skills/` directory, read the `SKILL.md` of any skill you need, and copy it into your own agent or CLI environment.

### 2. Use it as a shared team repository

Team members can contribute mature personal skills here, review them through pull requests, and evolve the repository into a maintained internal skill library.

## How to Contribute

1. Fork this repository
2. Create a branch
3. Add or update a skill under `skills/<skill-name>/`
4. Make sure the skill includes at least one `SKILL.md`
5. Only include files directly related to that skill
6. Open a Pull Request

See [CONTRIBUTING.md](CONTRIBUTING.md) for contribution details.

## Writing Guidelines

A good skill should clearly describe:

- what problem it solves
- when it should be triggered
- what inputs it needs
- how the workflow runs
- what constraints or risks it has
- what output it produces

Avoid vague descriptions. The goal is to make each skill directly reusable by others.

## Naming and Maintenance Rules

- Use `kebab-case` for skill directory names
- One directory should represent one skill
- Keep only directly related files
- Avoid committing unnecessary binary or temporary files
- Prefer compatibility-focused updates over unrelated refactors

## License

This repository is licensed under the [MIT License](LICENSE).
