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
git clone https://github.com/DWDyck/enterprise-slides-agent
cd enterprise-slides-agent
.\install.ps1
```

Requires Python 3.12+ (`winget install Python.Python.3.12 --scope user`). Packages: `python-pptx`, `lxml`, `Pillow`.
Restart Copilot so it discovers the agent.

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
python tests\smoke_test.py
```
