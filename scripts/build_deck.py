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
    "stats",
    "timeline",
    "comparison",
    "cards",
}

LAYOUT_ALIASES = {
    "kpi": "stats",
    "process": "timeline",
    "process_steps": "timeline",
    "card": "cards",
}


def rgb(hexstr: str) -> RGBColor:
    hexstr = str(hexstr).lstrip("#")
    return RGBColor(int(hexstr[0:2], 16), int(hexstr[2:4], 16), int(hexstr[4:6], 16))


class DeckBuilder:
    def __init__(self, spec: dict):
        theme_spec = spec.get("theme", "clean-light")
        palette_override = spec.get("palette")
        if isinstance(theme_spec, dict):
            palette_override = {**(theme_spec.get("palette") or {}), **(palette_override or {})}
            theme_spec = theme_spec.get("name") or theme_spec.get("base") or "clean-light"

        base_name = "clean-light" if theme_spec == "custom" else theme_spec
        if base_name not in THEMES:
            raise ValueError(
                f"unknown theme {theme_spec!r}; choose from {sorted(THEMES)} or 'custom'"
            )
        self.theme = dict(THEMES[base_name])
        if isinstance(palette_override, dict):
            for key, value in palette_override.items():
                if key in self.theme and isinstance(value, str):
                    self.theme[key] = value.lstrip("#")

        self.theme_name = str(theme_spec)
        if isinstance(palette_override, dict) and palette_override:
            self.theme_name = f"{self.theme_name} (palette override)"

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
        self._slide_no = 0

    # -- primitives -----------------------------------------------------

    def _new_slide(self) -> "object":
        slide = self.prs.slides.add_slide(self._blank)
        self._fill_background(slide)
        self._slide_no += 1
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

    def _card(self, slide, x, y, w, h, color):
        shp = self._rect(slide, x, y, w, h, color, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        try:
            shp.adjustments[0] = 0.05
        except Exception:
            pass
        return shp

    def _style_run(self, run, size, color, bold=False, italic=False, font=None):
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.italic = italic
        run.font.name = font or self.theme["body_font"]
        run.font.color.rgb = rgb(color)

    def _heading(self, slide, text, size, color, x=None, y=0.45, w=None, h=1.0):
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
            p.space_after = Pt(8)
        return box

    def _add_footer(self, slide) -> None:
        footer_color = self.theme["muted"]
        for shape in slide.shapes:
            try:
                x, y, w, h = shape.left or 0, shape.top or 0, shape.width or 0, shape.height or 0
            except Exception:
                continue
            if w >= self.prs.slide_width * 0.9 and h >= self.prs.slide_height * 0.9:
                footer_color = self.theme["band_text"]
        y = self.height_in - 0.42
        left_w = self.content_w - 1.1
        box, tf = self._textbox(slide, self.margin, y, left_w, 0.36)
        run = tf.paragraphs[0].add_run()
        run.text = str(self.deck_title or "")
        self._style_run(run, self.sizes["caption"], footer_color)
        box, tf = self._textbox(slide, self.width_in - self.margin - 1.0, y, 1.0, 0.36)
        tf.paragraphs[0].alignment = PP_ALIGN.RIGHT
        run = tf.paragraphs[0].add_run()
        run.text = str(self._slide_no)
        self._style_run(run, self.sizes["caption"], footer_color)

    def _normalize_items(self, raw):
        items = []
        for it in (raw or []):
            if isinstance(it, dict):
                items.append(it)
            elif isinstance(it, (list, tuple)):
                items.append({
                    "value": it[0] if len(it) > 0 else "",
                    "label": it[1] if len(it) > 1 else "",
                    "delta": it[2] if len(it) > 2 else None,
                    "title": it[0] if len(it) > 0 else "",
                    "body": it[1] if len(it) > 1 else "",
                    "items": it[2] if len(it) > 2 else None,
                })
            else:
                items.append({"value": str(it), "label": "", "title": str(it), "body": ""})
        return items

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
            _, stf = self._textbox(slide, self.margin, band_h + 0.5, self.content_w, 1.0)
            run = stf.paragraphs[0].add_run()
            run.text = sub
            self._style_run(run, self.sizes["body"] + 4, self.theme["muted"])
        self._notes(slide, slide_spec.get("notes"))

    def add_section(self, slide_spec):
        slide = self._new_slide()
        self._rect(slide, 0, self.height_in / 2 - 0.05, self.width_in, 0.1, self.theme["accent"])
        _, tf = self._textbox(slide, self.margin, self.height_in / 2 - 1.5, self.content_w, 1.4)
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
            slide, items, self.margin, 1.6, self.content_w, self.height_in - 2.2,
            self.sizes["body"], self.theme["text"],
        )
        if slide_spec.get("subtitle"):
            _, tf = self._textbox(slide, self.margin, 1.25, self.content_w, 0.4)
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
            self._card(slide, x, 1.6, col_w, self.height_in - 2.2, self.theme["surface"])
            self._add_bullets(
                slide, items, x + 0.3, 1.9, col_w - 0.6, self.height_in - 2.8,
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
            str(path), Inches(self.margin), Inches(1.6), width=Inches(self.content_w)
        )
        if pic.height > Inches(self.height_in - 2.3):
            ratio = (self.height_in - 2.3) / (pic.height / 914400)
            pic.width = Emu(int(pic.width * ratio))
            pic.height = Emu(int(pic.height * ratio))
            pic.left = Emu(int((self.prs.slide_width - pic.width) / 2))
            pic.top = Inches(1.6)
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
        n_cols = len(headers) or (len(rows[0]) if rows and isinstance(rows[0], (list, tuple)) else 0)
        if n_cols == 0 and rows:
            n_cols = 1
        n_rows = len(rows) + (1 if headers or rows else 0)
        if n_cols == 0:
            raise ValueError("table layout requires headers or rows")
        if n_rows == 0:
            n_rows = 1
        shape = slide.shapes.add_table(
            n_rows, n_cols, Inches(self.margin), Inches(1.75),
            Inches(self.content_w), Inches(self.height_in - 2.5),
        )
        table = shape.table
        if headers:
            for c, text in enumerate(headers):
                if c >= n_cols:
                    break
                cell = table.cell(0, c)
                cell.text = "" if text is None else str(text)
                cell.fill.solid()
                cell.fill.fore_color.rgb = rgb(self.theme["band"])
                for p in cell.text_frame.paragraphs:
                    for r in p.runs:
                        r.font.size = Pt(self.sizes["body"] - 3)
                        r.font.bold = True
                        r.font.color.rgb = rgb(self.theme["band_text"])
                        r.font.name = self.theme["body_font"]
        start_row = 1 if headers else 0
        for offset, row in enumerate(rows):
            r_i = start_row + offset
            if r_i >= n_rows:
                break
            row = row if isinstance(row, (list, tuple)) else [row]
            for c_i in range(n_cols):
                cell = table.cell(r_i, c_i)
                val = row[c_i] if c_i < len(row) else ""
                if val is None:
                    val = ""
                cell.text = str(val)
                if (r_i % 2 == 0) and headers:
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
        series_list = chart.get("series", [])
        for series in series_list:
            data.add_series(str(series.get("name", "Series")), tuple(series.get("values", [])))
        if not series_list:
            raise ValueError("chart layout requires at least one series")
        if ctype == "pie" and len(series_list) != 1:
            raise ValueError("pie charts require exactly one series")
        graphic = slide.shapes.add_chart(
            types[ctype], Inches(0.9), Inches(1.7), Inches(self.content_w - 0.6),
            Inches(self.height_in - 2.6), data,
        )
        ch = graphic.chart
        ch.has_legend = len(series_list) > 2
        if ctype == "pie":
            ch.plots[0].has_data_labels = True
        else:
            try:
                total_points = len(data.categories) * len(series_list)
                ch.plots[0].has_data_labels = total_points <= 12
            except Exception:
                pass
        for i, s in enumerate(ch.series):
            color = self.theme["accent"] if i == 0 else self.theme["muted"]
            try:
                s.format.fill.solid()
                s.format.fill.fore_color.rgb = rgb(color)
            except Exception:
                pass
            try:
                s.format.line.color.rgb = rgb(color)
            except Exception:
                pass
        self._notes(slide, slide_spec.get("notes"))

    def add_closing(self, slide_spec):
        slide = self._new_slide()
        band_h = self.height_in
        self._rect(slide, 0, 0, self.width_in, band_h, self.theme["band"])
        _, tf = self._textbox(
            slide, self.margin, self.height_in / 2 - 1.5, self.content_w, 1.4
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

    def add_stats(self, slide_spec):
        slide = self._new_slide()
        self._heading(slide, slide_spec.get("title", ""), self.sizes["title"], self.theme["text"])
        items = self._normalize_items(slide_spec.get("stats", slide_spec.get("items", [])))
        n = max(1, len(items))
        gap = 0.25
        card_w = (self.content_w - gap * (n - 1)) / n
        y = 1.6
        h = self.height_in - 2.4
        for i, it in enumerate(items):
            x = self.margin + i * (card_w + gap)
            self._card(slide, x, y, card_w, h, self.theme["surface"])
            _, tf = self._textbox(slide, x + 0.2, y + 0.5, card_w - 0.4, 1.3)
            tf.paragraphs[0].alignment = PP_ALIGN.CENTER
            run = tf.paragraphs[0].add_run()
            run.text = str(it.get("value", ""))
            self._style_run(run, self.sizes["deck_title"], self.theme["accent"], bold=True,
                            font=self.theme["title_font"])
            _, tf = self._textbox(slide, x + 0.2, y + 1.9, card_w - 0.4, 0.7)
            tf.paragraphs[0].alignment = PP_ALIGN.CENTER
            run = tf.paragraphs[0].add_run()
            run.text = str(it.get("label", ""))
            self._style_run(run, self.sizes["body"], self.theme["text"])
            delta = it.get("delta")
            if delta is not None and str(delta).strip() != "":
                _, tf = self._textbox(slide, x + 0.2, y + 2.6, card_w - 0.4, 0.5)
                tf.paragraphs[0].alignment = PP_ALIGN.CENTER
                run = tf.paragraphs[0].add_run()
                run.text = str(delta)
                color = self.theme["muted"]
                ds = str(delta).strip()
                if ds.startswith("+"):
                    color = self.theme["accent2"]
                elif ds.startswith("-"):
                    color = self.theme["accent3"]
                self._style_run(run, self.sizes["caption"] + 6, color, bold=True)
        self._notes(slide, slide_spec.get("notes"))

    def add_timeline(self, slide_spec):
        slide = self._new_slide()
        self._heading(slide, slide_spec.get("title", ""), self.sizes["title"], self.theme["text"])
        raw = slide_spec.get("steps", slide_spec.get("items", []))
        items = self._normalize_items(raw)
        # allow plain strings
        norm = []
        for it in raw or []:
            if isinstance(it, str):
                norm.append({"title": it, "body": ""})
            elif isinstance(it, dict):
                norm.append(it)
            elif isinstance(it, (list, tuple)):
                norm.append({"title": it[0] if len(it) > 0 else "", "body": it[1] if len(it) > 1 else ""})
            else:
                norm.append({"title": str(it), "body": ""})
        items = norm
        n = max(1, len(items))
        gap = 0.2
        card_w = (self.content_w - gap * (n - 1)) / n
        y = 1.7
        card_h = self.height_in - 2.5
        if n > 1:
            self._rect(slide, self.margin + 0.5, y + 0.45, self.content_w - 1.0, 0.06, self.theme["muted"])
        for i, it in enumerate(items):
            x = self.margin + i * (card_w + gap)
            circle = 0.6
            cx = x + (card_w - circle) / 2
            self._rect(slide, cx, y + 0.15, circle, circle, self.theme["accent"], shape=MSO_SHAPE.OVAL)
            _, tf = self._textbox(slide, cx, y + 0.15, circle, circle)
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            tf.paragraphs[0].alignment = PP_ALIGN.CENTER
            run = tf.paragraphs[0].add_run()
            run.text = str(i + 1)
            self._style_run(run, self.sizes["body"], self.theme["band_text"], bold=True,
                            font=self.theme["title_font"])
            _, tf = self._textbox(slide, x + 0.15, y + 0.9, card_w - 0.3, 0.6)
            tf.paragraphs[0].alignment = PP_ALIGN.CENTER
            run = tf.paragraphs[0].add_run()
            run.text = str(it.get("title", ""))
            self._style_run(run, self.sizes["body"], self.theme["text"], bold=True)
            _, tf = self._textbox(slide, x + 0.15, y + 1.55, card_w - 0.3, card_h - 1.55)
            tf.paragraphs[0].alignment = PP_ALIGN.CENTER
            run = tf.paragraphs[0].add_run()
            run.text = str(it.get("body", it.get("description", "")))
            self._style_run(run, self.sizes["caption"] + 2, self.theme["muted"])
        self._notes(slide, slide_spec.get("notes"))

    def add_comparison(self, slide_spec):
        slide = self._new_slide()
        self._heading(slide, slide_spec.get("title", ""), self.sizes["title"], self.theme["text"])
        headers = slide_spec.get("headers")
        columns = slide_spec.get("columns", slide_spec.get("items", []))
        if headers is None and columns and isinstance(columns[0], dict) and "header" in columns[0]:
            headers = [c.get("header", "") for c in columns]
            columns = [c.get("items", c.get("bullets", [])) for c in columns]
        headers = headers or []
        columns = columns or []
        n = max(1, min(3, max(len(headers), len(columns))))
        gap = 0.3
        col_w = (self.content_w - gap * (n - 1)) / n
        y = 1.6
        header_h = 0.5
        body_y = y + header_h + 0.15
        body_h = self.height_in - body_y - 0.7
        for i in range(n):
            x = self.margin + i * (col_w + gap)
            hdr = headers[i] if i < len(headers) else ""
            if str(hdr).strip() != "":
                self._rect(slide, x, y, col_w, header_h, self.theme["band"])
                _, tf = self._textbox(slide, x + 0.15, y, col_w - 0.3, header_h)
                tf.vertical_anchor = MSO_ANCHOR.MIDDLE
                tf.paragraphs[0].alignment = PP_ALIGN.CENTER
                run = tf.paragraphs[0].add_run()
                run.text = str(hdr)
                self._style_run(run, self.sizes["body"], self.theme["band_text"], bold=True)
            self._card(slide, x, body_y, col_w, body_h, self.theme["surface"])
            items = columns[i] if i < len(columns) else []
            if isinstance(items, dict):
                items = items.get("items", items.get("bullets", []))
            self._add_bullets(
                slide, items or [], x + 0.25, body_y + 0.2, col_w - 0.5, body_h - 0.4,
                self.sizes["body"] - 1, self.theme["text"],
            )
        self._notes(slide, slide_spec.get("notes"))

    def add_cards(self, slide_spec):
        slide = self._new_slide()
        self._heading(slide, slide_spec.get("title", ""), self.sizes["title"], self.theme["text"])
        raw = slide_spec.get("cards", slide_spec.get("items", []))
        items = []
        for it in raw or []:
            if isinstance(it, str):
                items.append({"title": it, "body": "", "bullets": []})
            elif isinstance(it, dict):
                items.append(it)
            elif isinstance(it, (list, tuple)):
                items.append({"title": it[0] if len(it) > 0 else "", "body": it[1] if len(it) > 1 else "", "bullets": it[2] if len(it) > 2 else []})
            else:
                items.append({"title": str(it), "body": "", "bullets": []})
        n = max(1, len(items))
        gap = 0.3
        card_w = (self.content_w - gap * (n - 1)) / n
        y = 1.6
        h = self.height_in - 2.4
        for i, it in enumerate(items):
            x = self.margin + i * (card_w + gap)
            self._card(slide, x, y, card_w, h, self.theme["surface"])
            _, tf = self._textbox(slide, x + 0.25, y + 0.25, card_w - 0.5, 0.6)
            run = tf.paragraphs[0].add_run()
            run.text = str(it.get("title", it.get("heading", "")))
            self._style_run(run, self.sizes["body"] + 1, self.theme["text"], bold=True)
            body = it.get("body", it.get("text", ""))
            bullets = it.get("bullets", [])
            if body:
                _, tf = self._textbox(slide, x + 0.25, y + 0.95, card_w - 0.5, 1.2)
                run = tf.paragraphs[0].add_run()
                run.text = str(body)
                self._style_run(run, self.sizes["body"] - 1, self.theme["muted"])
            if bullets:
                self._add_bullets(
                    slide, bullets, x + 0.25, y + 2.2, card_w - 0.5, h - 2.3,
                    self.sizes["body"] - 2, self.theme["text"],
                )
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
            "stats": self.add_stats,
            "timeline": self.add_timeline,
            "comparison": self.add_comparison,
            "cards": self.add_cards,
        }
        slides = spec.get("slides", [])
        if not slides:
            raise ValueError("deck spec has no slides")
        for slide_spec in slides:
            layout = slide_spec.get("layout", "bullets")
            layout = LAYOUT_ALIASES.get(layout, layout)
            if layout not in handlers:
                raise ValueError(
                    f"unknown layout {layout!r}; choose from {sorted(LAYOUTS)}"
                )
            handlers[layout](slide_spec)
            self._add_footer(self.prs.slides[-1])


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
