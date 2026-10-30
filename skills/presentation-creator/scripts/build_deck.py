#!/usr/bin/env python3
"""Build a themed .pptx deck from a JSON deck spec.

Usage:
    python3 build_deck.py deck.json -o deck.pptx [--force]

The JSON spec contract is documented in the skill's SKILL.md. Output is
deterministic for a given spec, theme, and python-pptx version.
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

try:
    from pptx import Presentation
    from pptx.chart.data import CategoryChartData
    from pptx.dml.color import RGBColor
    from pptx.enum.chart import XL_CHART_TYPE
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
    from pptx.util import Emu, Inches, Pt
except ModuleNotFoundError:  # pragma: no cover - environment guard
    sys.stderr.write(
        "error: python-pptx is not installed. Install it with\n"
        "  python3 -m venv /tmp/pptx-venv && /tmp/pptx-venv/bin/pip install python-pptx\n"
        "or `pip install --user python-pptx`.\n"
    )
    raise SystemExit(2)


THEMES: dict[str, dict] = {
    "clean-light": {
        "bg": "FFFFFF",
        "surface": "F2F4F7",
        "text": "1A1A1A",
        "muted": "5A6472",
        "accent": "1F6FEB",
        "accent2": "0E9F6E",
        "accent3": "D9381E",
        "band": "1F6FEB",
        "band_text": "FFFFFF",
        "title_font": "Segoe UI Semibold",
        "body_font": "Segoe UI",
    },
    "corporate-blue": {
        "bg": "FFFFFF",
        "surface": "E8EEF7",
        "text": "0B1F33",
        "muted": "4A5A6A",
        "accent": "0B3D91",
        "accent2": "2E7DD1",
        "accent3": "C9A227",
        "band": "0B3D91",
        "band_text": "FFFFFF",
        "title_font": "Calibri",
        "body_font": "Calibri",
    },
    "dark-tech": {
        "bg": "0D1117",
        "surface": "161B22",
        "text": "E6EDF3",
        "muted": "8B949E",
        "accent": "58A6FF",
        "accent2": "3FB950",
        "accent3": "F85149",
        "band": "0B3D91",
        "band_text": "E6EDF3",
        "title_font": "Segoe UI Semibold",
        "body_font": "Segoe UI",
    },
}

ASPECTS = {
    "16:9": (13.333, 7.5),
    "4:3": (10.0, 7.5),
}

LAYOUTS = {
    "title",
    "section",
    "bullets",
    "two_column",
    "image",
    "table",
    "quote",
    "chart",
    "closing",
}


def rgb(hexstr: str) -> RGBColor:
    hexstr = hexstr.lstrip("#")
    return RGBColor(int(hexstr[0:2], 16), int(hexstr[2:4], 16), int(hexstr[4:6], 16))


class DeckBuilder:
    def __init__(self, spec: dict):
        theme_name = spec.get("theme", "clean-light")
        if theme_name not in THEMES:
            raise ValueError(
                f"unknown theme {theme_name!r}; choose from {sorted(THEMES)}"
            )
        self.theme = THEMES[theme_name]
        self.theme_name = theme_name

        aspect = spec.get("aspect", "16:9")
        if aspect not in ASPECTS:
            raise ValueError(f"unknown aspect {aspect!r}; choose from {sorted(ASPECTS)}")
        width_in, height_in = ASPECTS[aspect]

        self.prs = Presentation()
        self.prs.slide_width = Inches(width_in)
        self.prs.slide_height = Inches(height_in)
        self.width_in = width_in
        self.height_in = height_in
        self.margin = 0.6
        self.content_w = width_in - 2 * self.margin

        scale = 1.0 if aspect == "16:9" else 0.93
        self.sizes = {
            "deck_title": round(40 * scale),
            "title": round(30 * scale),
            "section": round(34 * scale),
            "body": round(18 * scale),
            "caption": 12,
            "quote": round(26 * scale),
        }
        self._blank = self.prs.slide_layouts[6]
        self.deck_title = spec.get("title")
        self.deck_subtitle = spec.get("subtitle")

    # -- primitives -----------------------------------------------------

    def _new_slide(self) -> "object":
        slide = self.prs.slides.add_slide(self._blank)
        self._fill_background(slide)
        return slide

    def _fill_background(self, slide) -> None:
        try:
            fill = slide.background.fill
            fill.solid()
            fill.fore_color.rgb = rgb(self.theme["bg"])
        except Exception:
            band = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE, 0, 0, self.prs.slide_width, self.prs.slide_height
            )
            band.fill.solid()
            band.fill.fore_color.rgb = rgb(self.theme["bg"])
            band.line.fill.background()
            band.shadow.inherit = False

    def _textbox(self, slide, x, y, w, h):
        box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = box.text_frame
        tf.word_wrap = True
        return box, tf

    def _rect(self, slide, x, y, w, h, color, shape=MSO_SHAPE.RECTANGLE):
        shp = slide.shapes.add_shape(
            shape, Inches(x), Inches(y), Inches(w), Inches(h)
        )
        shp.fill.solid()
        shp.fill.fore_color.rgb = rgb(color)
        shp.line.fill.background()
        shp.shadow.inherit = False
        return shp

    def _style_run(self, run, size, color, bold=False, italic=False, font=None):
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.italic = italic
        run.font.name = font or self.theme["body_font"]
        run.font.color.rgb = rgb(color)

    def _heading(self, slide, text, size, color, x=None, y=0.45, w=None, h=1.1):
        x = self.margin if x is None else x
        w = self.content_w if w is None else w
        box, tf = self._textbox(slide, x, y, w, h)
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.LEFT
        run = p.add_run()
        run.text = text or ""
        self._style_run(
            run, size, color, bold=True, font=self.theme["title_font"]
        )
        return box

    def _notes(self, slide, notes) -> None:
        if notes:
            slide.notes_slide.notes_text_frame.text = str(notes)

    def _add_bullets(self, slide, items, x, y, w, h, size, color):
        box, tf = self._textbox(slide, x, y, w, h)
        tf.vertical_anchor = MSO_ANCHOR.TOP
        for i, text in enumerate(items):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.level = 0
            run = p.add_run()
            run.text = f"\u2022  {text}"
            self._style_run(run, size, color)
            p.space_after = Pt(10)
        return box

    # -- layouts --------------------------------------------------------

    def add_title(self, slide_spec):
        slide = self._new_slide()
        band_h = 2.4
        self._rect(slide, 0, 0, self.width_in, band_h, self.theme["band"])
        _, tf = self._textbox(slide, self.margin, 0.9, self.content_w, 1.2)
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        run = tf.paragraphs[0].add_run()
        run.text = slide_spec.get("title") or self.deck_title or ""
        self._style_run(
            run, self.sizes["deck_title"], self.theme["band_text"], bold=True,
            font=self.theme["title_font"],
        )
        sub = slide_spec.get("subtitle") or self.deck_subtitle
        if sub:
            _, stf = self._textbox(slide, self.margin, band_h + 0.6, self.content_w, 1.0)
            run = stf.paragraphs[0].add_run()
            run.text = sub
            self._style_run(run, self.sizes["body"] + 4, self.theme["muted"])
        self._notes(slide, slide_spec.get("notes"))

    def add_section(self, slide_spec):
        slide = self._new_slide()
        self._rect(slide, 0, self.height_in / 2 - 0.05, self.width_in, 0.1, self.theme["accent"])
        _, tf = self._textbox(slide, self.margin, self.height_in / 2 - 1.6, self.content_w, 1.5)
        tf.vertical_anchor = MSO_ANCHOR.BOTTOM
        run = tf.paragraphs[0].add_run()
        run.text = slide_spec.get("title", "")
        self._style_run(
            run, self.sizes["section"], self.theme["text"], bold=True,
            font=self.theme["title_font"],
        )
        sub = slide_spec.get("subtitle")
        if sub:
            _, stf = self._textbox(slide, self.margin, self.height_in / 2 + 0.2, self.content_w, 1.0)
            run = stf.paragraphs[0].add_run()
            run.text = sub
            self._style_run(run, self.sizes["body"], self.theme["muted"])
        self._notes(slide, slide_spec.get("notes"))

    def add_bullets(self, slide_spec):
        slide = self._new_slide()
        self._heading(slide, slide_spec.get("title", ""), self.sizes["title"], self.theme["text"])
        items = slide_spec.get("bullets", [])
        self._add_bullets(
            slide, items, self.margin, 1.7, self.content_w, self.height_in - 2.3,
            self.sizes["body"], self.theme["text"],
        )
        if slide_spec.get("subtitle"):
            _, tf = self._textbox(slide, self.margin, 1.35, self.content_w, 0.4)
            run = tf.paragraphs[0].add_run()
            run.text = slide_spec["subtitle"]
            self._style_run(run, self.sizes["caption"] + 2, self.theme["muted"])
        self._notes(slide, slide_spec.get("notes"))

    def add_two_column(self, slide_spec):
        slide = self._new_slide()
        self._heading(slide, slide_spec.get("title", ""), self.sizes["title"], self.theme["text"])
        gap = 0.4
        col_w = (self.content_w - gap) / 2
        columns = slide_spec.get("columns", [[], []])
        left = columns[0] if len(columns) > 0 else []
        right = columns[1] if len(columns) > 1 else []
        for idx, items in enumerate((left, right)):
            x = self.margin + idx * (col_w + gap)
            self._rect(slide, x, 1.7, col_w, self.height_in - 2.3, self.theme["surface"])
            self._add_bullets(
                slide, items, x + 0.3, 2.0, col_w - 0.6, self.height_in - 2.9,
                self.sizes["body"] - 1, self.theme["text"],
            )
        self._notes(slide, slide_spec.get("notes"))

    def add_image(self, slide_spec):
        slide = self._new_slide()
        self._heading(slide, slide_spec.get("title", ""), self.sizes["title"], self.theme["text"])
        img = slide_spec.get("image")
        if not img:
            raise ValueError("image layout requires an 'image' path")
        path = Path(img)
        if not path.is_file():
            raise FileNotFoundError(f"image not found: {img}")
        pic = slide.shapes.add_picture(
            str(path), Inches(self.margin), Inches(1.7), width=Inches(self.content_w)
        )
        if pic.height > Inches(self.height_in - 2.3):
            ratio = (self.height_in - 2.3) / (pic.height / 914400)
            pic.width = Emu(int(pic.width * ratio))
            pic.height = Emu(int(pic.height * ratio))
            pic.left = Emu(int((self.prs.slide_width - pic.width) / 2))
            pic.top = Inches(1.7)
        caption = slide_spec.get("caption")
        if caption:
            _, tf = self._textbox(
                slide, self.margin, self.height_in - 0.7, self.content_w, 0.4
            )
            run = tf.paragraphs[0].add_run()
            run.text = caption
            self._style_run(run, self.sizes["caption"], self.theme["muted"])
        self._notes(slide, slide_spec.get("notes"))

    def add_table(self, slide_spec):
        slide = self._new_slide()
        self._heading(slide, slide_spec.get("title", ""), self.sizes["title"], self.theme["text"])
        tbl = slide_spec.get("table", {})
        headers = tbl.get("headers", [])
        rows = tbl.get("rows", [])
        n_cols = len(headers) or (len(rows[0]) if rows else 0)
        n_rows = len(rows) + 1
        if n_cols == 0:
            raise ValueError("table layout requires headers or rows")
        shape = slide.shapes.add_table(
            n_rows, n_cols, Inches(self.margin), Inches(1.8),
            Inches(self.content_w), Inches(self.height_in - 2.6),
        )
        table = shape.table
        for c, text in enumerate(headers):
            cell = table.cell(0, c)
            cell.text = str(text)
            cell.fill.solid()
            cell.fill.fore_color.rgb = rgb(self.theme["band"])
            for p in cell.text_frame.paragraphs:
                for r in p.runs:
                    r.font.size = Pt(self.sizes["body"] - 3)
                    r.font.bold = True
                    r.font.color.rgb = rgb(self.theme["band_text"])
                    r.font.name = self.theme["body_font"]
        for r_i, row in enumerate(rows, start=1):
            for c_i in range(n_cols):
                cell = table.cell(r_i, c_i)
                cell.text = str(row[c_i]) if c_i < len(row) else ""
                if r_i % 2 == 0:
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = rgb(self.theme["surface"])
                for p in cell.text_frame.paragraphs:
                    for r in p.runs:
                        r.font.size = Pt(self.sizes["body"] - 3)
                        r.font.color.rgb = rgb(self.theme["text"])
                        r.font.name = self.theme["body_font"]
        self._notes(slide, slide_spec.get("notes"))

    def add_quote(self, slide_spec):
        slide = self._new_slide()
        self._rect(slide, self.margin, 2.0, 0.12, 3.0, self.theme["accent"])
        box, tf = self._textbox(slide, self.margin + 0.5, 2.0, self.content_w - 0.5, 2.4)
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        run = tf.paragraphs[0].add_run()
        run.text = f"\u201c{slide_spec.get('quote', '')}\u201d"
        self._style_run(
            run, self.sizes["quote"], self.theme["text"], italic=True,
            font=self.theme["title_font"],
        )
        attr = slide_spec.get("attribution")
        if attr:
            p = tf.add_paragraph()
            p.space_before = Pt(12)
            run = p.add_run()
            run.text = f"\u2014 {attr}"
            self._style_run(run, self.sizes["caption"] + 2, self.theme["muted"])
        self._notes(slide, slide_spec.get("notes"))

    def add_chart(self, slide_spec):
        slide = self._new_slide()
        self._heading(slide, slide_spec.get("title", ""), self.sizes["title"], self.theme["text"])
        chart = slide_spec.get("chart", {})
        ctype = chart.get("type", "bar")
        types = {
            "bar": XL_CHART_TYPE.COLUMN_CLUSTERED,
            "line": XL_CHART_TYPE.LINE_MARKERS,
            "pie": XL_CHART_TYPE.PIE,
        }
        if ctype not in types:
            raise ValueError(f"unknown chart type {ctype!r}; choose from {sorted(types)}")
        data = CategoryChartData()
        data.categories = [str(c) for c in chart.get("categories", [])]
        for series in chart.get("series", []):
            data.add_series(str(series.get("name", "Series")), tuple(series.get("values", [])))
        if not chart.get("series"):
            raise ValueError("chart layout requires at least one series")
        if ctype == "pie" and len(chart["series"]) != 1:
            raise ValueError("pie charts require exactly one series")
        graphic = slide.shapes.add_chart(
            types[ctype], Inches(0.9), Inches(1.8), Inches(self.content_w - 0.6),
            Inches(self.height_in - 2.7), data,
        )
        graphic.chart.has_legend = ctype != "pie"
        if ctype == "pie":
            graphic.chart.plots[0].has_data_labels = True
        self._notes(slide, slide_spec.get("notes"))

    def add_closing(self, slide_spec):
        slide = self._new_slide()
        band_h = self.height_in
        self._rect(slide, 0, 0, self.width_in, band_h, self.theme["band"])
        _, tf = self._textbox(
            slide, self.margin, self.height_in / 2 - 1.6, self.content_w, 1.4
        )
        tf.vertical_anchor = MSO_ANCHOR.BOTTOM
        run = tf.paragraphs[0].add_run()
        run.text = slide_spec.get("title", "Thank you")
        self._style_run(
            run, self.sizes["deck_title"] - 6, self.theme["band_text"], bold=True,
            font=self.theme["title_font"],
        )
        sub = slide_spec.get("subtitle") or slide_spec.get("quote")
        if sub:
            _, stf = self._textbox(
                slide, self.margin, self.height_in / 2 + 0.2, self.content_w, 1.0
            )
            run = stf.paragraphs[0].add_run()
            run.text = sub
            self._style_run(run, self.sizes["body"] + 2, self.theme["band_text"])
        self._notes(slide, slide_spec.get("notes"))

    def build(self, spec: dict) -> None:
        handlers = {
            "title": self.add_title,
            "section": self.add_section,
            "bullets": self.add_bullets,
            "two_column": self.add_two_column,
            "image": self.add_image,
            "table": self.add_table,
            "quote": self.add_quote,
            "chart": self.add_chart,
            "closing": self.add_closing,
        }
        slides = spec.get("slides", [])
        if not slides:
            raise ValueError("deck spec has no slides")
        for slide_spec in slides:
            layout = slide_spec.get("layout", "bullets")
            if layout not in handlers:
                raise ValueError(
                    f"unknown layout {layout!r}; choose from {sorted(LAYOUTS)}"
                )
            handlers[layout](slide_spec)


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="Build a .pptx deck from a JSON spec.")
    parser.add_argument("spec", help="path to the JSON deck spec")
    parser.add_argument("-o", "--output", required=True, help="output .pptx path")
    parser.add_argument("--force", action="store_true", help="overwrite an existing file")
    args = parser.parse_args(argv)

    spec_path = Path(args.spec)
    if not spec_path.is_file():
        print(f"error: spec not found: {spec_path}", file=sys.stderr)
        return 2
    out_path = Path(args.output)
    if out_path.exists() and not args.force:
        print(
            f"error: {out_path} exists; pass --force to overwrite",
            file=sys.stderr,
        )
        return 2

    try:
        spec = json.loads(spec_path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        print(f"error: invalid JSON in {spec_path}: {exc}", file=sys.stderr)
        return 2

    try:
        builder = DeckBuilder(spec)
        builder.build(spec)
        out_path.parent.mkdir(parents=True, exist_ok=True)
        builder.prs.save(str(out_path))
    except (ValueError, FileNotFoundError) as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 2

    n = len(builder.prs.slides._sldIdLst)  # noqa: SLF001 - documented slide count
    print(f"Wrote {n} slides to {out_path} (theme: {builder.theme_name}, {spec.get('aspect', '16:9')})")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
