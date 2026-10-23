# enterprise-slides agent

A GitHub Copilot custom agent that builds **editable, enterprise-ready PowerPoint decks** in a hand-drawn,
"Squiggly Careers"-inspired style (squiggle underlines, circled numbers, arrows, journey paths), starting
from your **pre-approved `.pptx`/`.potx` templates**.

## What's in the repo

| Path | Purpose |
|---|---|
| `agents/enterprise-slides.agent.md` | The agent definition (workflow, style, enterprise rules, tooling choice) |
| `lib/slide_helpers.py` | Tested python-pptx helpers: open template safely, map layouts, native Bézier doodles, alt text, notes, structural audit |
| `tests/smoke_test.py` | End-to-end test: build from a template, verify template untouched, audit output |
| `style.example.json` | Copied to `~/.copilot/slide-assets/style.json` to override palette/fonts/sizes |
| `install.ps1` | Per-user installer (no admin rights) |

## Install (Windows)

```powershell
git clone https://github.com/GitDdyck/enterprise-slides-agent
cd enterprise-slides-agent
.\install.ps1
```

Requires Python 3.12+ from an approved software channel. The installer creates an isolated
per-user environment at `~/.copilot/slide-assets/venv`, installs the pinned packages in
`requirements.txt`, and runs the smoke test. It does not elevate privileges or install packages
machine-wide. If the package install or test fails, the installer stops and reports the failure.
If your approved package source is configured through pip, the installer uses that configuration;
it does not try another source if access is blocked. Restart Copilot so it discovers the agent.

### Enterprise data handling

Before using real work material, confirm with your organization that the signed-in Copilot
account/tenant, selected model, and VS Code agent harness are approved for that material's
classification. If this is uncertain, use sanitized sample content and ask your IT/security team
before proceeding. The agent instructions are guidance, not an enforcement boundary: local
PowerPoint/Python processing does not by itself mean prompts or supplied content stay local.
Use only templates and assets approved for the selected environment; do not commit company
templates, generated decks, or confidential content to this repository.

## Customise

Put these in `~/.copilot/slide-assets/`:

- `templates/` — your corporate `.pptx`/`.potx` (the agent always builds from a copy; never edits the master)
- `doodles/` — your own SVG/PNG doodles
- `fonts/` — handwriting font(s) with licence notes
- `style.json` — palette, fonts, sizes

## Known limits

- Ink strokes, animations/transitions and SmartArt can't be authored programmatically; the agent leaves marked space for live stylus inking instead.
- Visual rendering QA needs desktop PowerPoint (COM) or LibreOffice; without them only structural checks (`audit`) run.
- Palette hex values are approximations of the Squiggly Careers look, not official brand values.

## Test

```powershell
& "$env:USERPROFILE\.copilot\slide-assets\venv\Scripts\python.exe" tests\smoke_test.py
```

The GitHub Actions smoke test uses a GitHub-hosted Windows runner. Enable or rely on that cloud
CI only if repository code is permitted to run in that environment; otherwise run the local test
with the approved Python environment above.
