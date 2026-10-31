# kilo-presentation-creator

A [Kilo](https://kilo.ai) skill that builds polished PowerPoint decks (`.pptx`) from a brief,
outline, or source document with [python-pptx](https://python-pptx.readthedocs.io/), then runs
an automated QA pass for overflow, contrast, and font floors before you deliver it.

This repository is the canonical source of the `presentation-creator` skill. It is also listed
in the [Kilo marketplace](https://github.com/Kilo-Org/kilo-marketplace).

## What it does

- **Workflow, not vibes.** Preflight, clarify, outline, theme, build, QA gate, optional
  preview, deliver — encoded in `SKILL.md`.
- **Deterministic builds.** A JSON deck spec drives `scripts/build_deck.py`; the same input
  produces the same `.pptx`. Custom slides use the verified snippets in
  `references/python-pptx-recipes.md`.
- **Three themes.** `clean-light`, `corporate-blue`, `dark-tech`, each with a palette, type
  scale, and grid (`references/design-system.md`).
- **Automated QA gate.** `scripts/validate_deck.py` checks slide bounds, estimated text
  overflow, font floors (body ≥ 14pt, titles ≥ 28pt), bullet counts, missing titles/notes,
  duplicate titles, and WCAG contrast. Errors block delivery.

## Install

Copy this directory (the repo root) into a Kilo skills location:

- Project-local: `.kilo/skills/presentation-creator/`
- Global: `~/.config/kilo/skills/presentation-creator/`

The skill lives at the repository root, so copy the whole repo (or its contents) into that directory.

Or install it from the Kilo marketplace.

The scripts need Python 3.9+ with `python-pptx`:

```sh
python3 -m venv /tmp/pptx-venv
/tmp/pptx-venv/bin/pip install python-pptx
python3 -c "import pptx; print(pptx.__version__)"   # 1.0.x tested
```

`matplotlib` and `Pillow` are optional (chart/image slides); LibreOffice (`soffice`) is
optional (PDF preview only).

## Usage

Author a deck spec (`examples/example-deck.json` is a complete sample) and build it:

```sh
python3 scripts/build_deck.py deck.json -o deck.pptx
python3 scripts/validate_deck.py deck.pptx
```

The validator exits `0` when the deck passes (warnings-only also exits `0`) and `1` when it
reports errors. Fix every error and re-run until it is clean.

In Kilo, just ask for slides — the skill activates on requests for a deck, presentation,
pitch, keynote, or slide outline (and on `*.pptx` files).

## Layout

```
.
├── SKILL.md
├── LICENSE
├── references/   python-pptx-recipes.md, design-system.md, troubleshooting.md
├── scripts/      build_deck.py, validate_deck.py, render_preview.py, test_deck.py
└── examples/     example-deck.json, layouts-deck.json, example-brief.md
```

## License

Apache-2.0. See [LICENSE](LICENSE).
