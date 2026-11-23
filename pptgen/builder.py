"""Turn a `Deck` into a real .pptx file with python-pptx.

Layouts: cover, bullets (with/without photo), cards, two_column, stats, chart, section.
Text is auto-sized to fill its box without overflowing.
"""
import io
import math
from pathlib import Path
from typing import Callable, List, Optional, Sequence, Tuple

from PIL import Image
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.util import Emu, Inches, Pt

from .images import fetch_image
from .models import ChartSpec, Deck, Point, SlideSpec, ThemeSpec

W, H = Inches(13.333), Inches(7.5)  # 16:9
MARGIN = Inches(0.7)
GAP = Inches(0.3)
CONTENT_BOTTOM = H - Inches(0.75)
PHOTO_X = 0.68  # bullet slides: photo strip starts at 68% of the width

CHART_TYPES = {
    "column": XL_CHART_TYPE.COLUMN_CLUSTERED,
    "bar": XL_CHART_TYPE.BAR_CLUSTERED,
    "line": XL_CHART_TYPE.LINE_MARKERS,
    "pie": XL_CHART_TYPE.PIE,
}

# Average character width as a fraction of font size (used to estimate wrapping).
# Measured from the real fonts (regular weight) + ~5% safety margin.
CHAR_WIDTH = {
    "Calibri": 0.43, "Arial": 0.47, "Segoe UI": 0.465, "Verdana": 0.535,
    "Trebuchet MS": 0.48, "Georgia": 0.47, "Cambria": 0.45, "Garamond": 0.41,
    "Century Gothic": 0.505,
}
DEFAULT_CHAR_WIDTH = 0.47
# Natural line height of each font as a multiple of the font size (measured).
LINE_HEIGHT = {
    "Calibri": 1.22, "Arial": 1.15, "Segoe UI": 1.33, "Verdana": 1.22,
    "Trebuchet MS": 1.16, "Georgia": 1.14, "Cambria": 1.17, "Garamond": 1.13,
    "Century Gothic": 1.23,
}
DEFAULT_LINE_HEIGHT = 1.2


# ---------------------------------------------------------------- colours
def _hex(c: Optional[str], fallback: str = "#333333") -> str:
    s = (c or "").strip().lstrip("#")[:6]
    try:
        int(s, 16)
        if len(s) == 6:
            return s.upper()
    except ValueError:
        pass
    return fallback.lstrip("#").upper()


def rgb(c: Optional[str], fallback: str = "#333333") -> RGBColor:
    return RGBColor.from_string(_hex(c, fallback))


def mix(a: str, b: str, t: float) -> str:
    """Blend colour a towards b by fraction t (0..1)."""
    x, y = _hex(a), _hex(b)
    return "#" + "".join(
        f"{round(int(x[i:i + 2], 16) * (1 - t) + int(y[i:i + 2], 16) * t):02X}" for i in (0, 2, 4))


def luminance(c: str) -> float:
    s = _hex(c)

    def chan(v: float) -> float:
        return v / 12.92 if v <= 0.03928 else ((v + 0.055) / 1.055) ** 2.4

    r, g, b = (chan(int(s[i:i + 2], 16) / 255) for i in (0, 2, 4))
    return 0.2126 * r + 0.7152 * g + 0.0722 * b


def contrast(a: str, b: str) -> float:
    la, lb = sorted((luminance(a), luminance(b)), reverse=True)
    return (la + 0.05) / (lb + 0.05)


def on_color(c: str) -> str:
    """Readable text colour (white or near-black) on top of colour c."""
    return "#FFFFFF" if contrast("#FFFFFF", c) >= contrast("#111111", c) else "#111111"


# ---------------------------------------------------------------- text fitting
def fit_size(items: Sequence[Tuple[str, float]], width, height, max_pt: int, min_pt: int,
             font: str, gap: float = 0.55, line: float = 1.12) -> int:
    """Largest font size at which all paragraphs fit in the box.

    items: (text, scale) pairs; scale lets a heading be bigger than body text.
    """
    for size in range(max_pt, min_pt - 1, -1):
        if fits(items, width, height, size, font, gap, line):
            return size
    return min_pt


def fits(items: Sequence[Tuple[str, float]], width, height, size: float, font: str,
         gap: float = 0.55, line: float = 1.12) -> bool:
    """Estimate whether the paragraphs fit in the box at this font size.

    line: paragraph line-spacing multiplier, applied on top of the font's own line height.
    """
    width_pt = Emu(int(width)).pt - 16   # text box internal margins
    height_pt = Emu(int(height)).pt - 10
    cw = CHAR_WIDTH.get(font, DEFAULT_CHAR_WIDTH)
    lh = LINE_HEIGHT.get(font, DEFAULT_LINE_HEIGHT) * line
    need = 0.0
    for text, scale in items:
        s = size * scale
        need += wrap_lines(text, width_pt, s * cw) * s * lh
    need += max(0, len(items) - 1) * size * gap
    return need <= height_pt


def wrap_lines(text: str, width_pt: float, char_pt: float) -> int:
    """Count lines by wrapping whole words, like PowerPoint does."""
    max_chars = max(1.0, width_pt / char_pt)
    lines, used = 1, 0.0
    for word in text.split():
        n = len(word)
        if used == 0:
            used = n
        elif used + 1 + n <= max_chars:
            used += 1 + n
        else:
            lines += 1
            used = n
        while used > max_chars:  # a single word longer than the line
            lines += 1
            used -= max_chars
    return lines


def shorten(text: str, max_chars: int) -> str:
    """Cut text at a word boundary and add an ellipsis (last-resort overflow guard)."""
    if len(text) <= max_chars:
        return text
    cut = text[:max_chars].rsplit(" ", 1)[0].rstrip(",;:-")
    return cut + "…"


def shrink_to_fit(texts: List[str], build_items, width, height, min_pt: int, font: str,
                  floor: int = 40) -> List[str]:
    """Shorten the longest texts until everything fits at the minimum font size."""
    texts = list(texts)
    for _ in range(12):
        if fits(build_items(texts), width, height, min_pt, font) or max(map(len, texts)) <= floor:
            break
        limit = int(max(map(len, texts)) * 0.85)
        texts = [shorten(t, limit) for t in texts]
    return texts


class DeckBuilder:
    def __init__(self, theme: ThemeSpec):
        t = self.t = theme
        self.bg = "#" + _hex(t.background, "#FFFFFF")
        # Guard against AI themes with poor contrast
        self.title_c = t.title_color if contrast(t.title_color, self.bg) >= 3 else on_color(self.bg)
        self.text_c = t.text_color if contrast(t.text_color, self.bg) >= 4 else on_color(self.bg)
        self.accent = "#" + _hex(t.accent, "#2563EB")
        self.accent2 = "#" + _hex(t.accent2, "#64748B")
        self.muted = mix(self.text_c, self.bg, 0.35)
        self.surface = mix(self.bg, self.text_c, 0.06)       # card background
        self.tint = mix(self.bg, self.accent, 0.12)          # callout background
        self.palette = [self.accent, self.accent2, "#F59E0B", "#10B981", "#EF4444", "#6366F1"]
        self.deck_title = ""

        self.prs = Presentation()
        self.prs.slide_width, self.prs.slide_height = W, H
        self.blank = self.prs.slide_layouts[6]

    # ============================================================ primitives
    def _new_slide(self):
        slide = self.prs.slides.add_slide(self.blank)
        fill = slide.background.fill
        fill.solid()
        fill.fore_color.rgb = rgb(self.bg)
        return slide

    def _rect(self, slide, x, y, w, h, color, rounded=False, radius=0.06):
        shape = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE,
            int(x), int(y), int(w), int(h))
        if rounded:
            shape.adjustments[0] = radius
        shape.fill.solid()
        shape.fill.fore_color.rgb = rgb(color)
        shape.line.fill.background()
        shape.shadow.inherit = False
        return shape

    def _frame(self, slide, x, y, w, h, anchor=MSO_ANCHOR.TOP):
        box = slide.shapes.add_textbox(int(x), int(y), int(w), int(h))
        tf = box.text_frame
        tf.word_wrap = True
        tf.auto_size = MSO_AUTO_SIZE.NONE
        tf.vertical_anchor = anchor
        return tf

    @staticmethod
    def _run(p, text, size, color, font, bold=False, italic=False):
        r = p.add_run()
        r.text = text
        f = r.font
        f.size, f.bold, f.italic, f.name = Pt(size), bold, italic, font
        f.color.rgb = rgb(color)
        return r

    def _text(self, slide, text, x, y, w, h, size, color, font=None, bold=False,
              italic=False, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
        tf = self._frame(slide, x, y, w, h, anchor)
        p = tf.paragraphs[0]
        p.alignment = align
        self._run(p, text, size, color, font or self.t.body_font, bold, italic)
        return tf

    def _picture_fill(self, slide, path: Path, x, y, w, h):
        """Insert an image cropped to exactly fill the box (no stretching)."""
        with Image.open(path) as im:
            iw, ih = im.size
        pic = slide.shapes.add_picture(str(path), int(x), int(y), int(w), int(h))
        box_ratio, img_ratio = w / h, iw / ih
        if img_ratio > box_ratio:
            excess = 1 - box_ratio / img_ratio
            pic.crop_left = pic.crop_right = excess / 2
        else:
            excess = 1 - img_ratio / box_ratio
            pic.crop_top = pic.crop_bottom = excess / 2

    @staticmethod
    def _notes(slide, notes, credit=None):
        text = "\n\n".join(t for t in (notes, credit) if t)
        if text:
            slide.notes_slide.notes_text_frame.text = text

    # ============================================================ building blocks
    def _header(self, slide, spec: SlideSpec, width) -> int:
        """Title + accent bar + optional subtitle. Returns the y where content starts."""
        font = self.t.title_font
        size = fit_size([(spec.title, 1.0)], width, Inches(1.0), 34, 24, font)
        self._text(slide, spec.title, MARGIN, Inches(0.4), width, Inches(1.0), size,
                   self.title_c, font, bold=True, anchor=MSO_ANCHOR.BOTTOM)
        self._rect(slide, MARGIN + Inches(0.08), Inches(1.45), Inches(1.0), Inches(0.06), self.accent)
        y = Inches(1.65)
        if spec.subtitle:
            sub = fit_size([(spec.subtitle, 1.0)], width, Inches(0.6), 18, 14, self.t.body_font)
            self._text(slide, spec.subtitle, MARGIN, y, width, Inches(0.6), sub,
                       self.muted, italic=True)
            y += Inches(0.65)
        return y + Inches(0.15)

    def _footer(self, slide, number: int, light=False):
        color = self.muted if not light else on_color(self.accent)
        self._text(slide, f"{number:02d}   |   {shorten(self.deck_title, 70)}", MARGIN, H - Inches(0.55),
                   Inches(7), Inches(0.35), 10, color)

    def _callout(self, slide, text, x, w, bottom=CONTENT_BOTTOM) -> int:
        """Accent-tinted insight box sitting on `bottom`. Returns its top y."""
        size = 15
        label = "Key insight:  "
        text = shorten(text, 220)
        cw = CHAR_WIDTH.get(self.t.body_font, DEFAULT_CHAR_WIDTH)
        per_line = max(1, int((Emu(int(w)).pt - 50) / (size * cw)))
        lines = min(3, math.ceil((len(label) + len(text)) * 1.08 / per_line))
        h = Pt(lines * size * 1.3 + 26)
        y = bottom - h
        self._rect(slide, x, y, w, h, self.tint, rounded=True, radius=0.12)
        self._rect(slide, x, y, Inches(0.09), h, self.accent)
        tf = self._frame(slide, x + Inches(0.25), y, w - Inches(0.35), h, MSO_ANCHOR.MIDDLE)
        p = tf.paragraphs[0]
        self._run(p, label, size, self.accent, self.t.body_font, bold=True)
        self._run(p, text, size, self.text_c, self.t.body_font)
        return y - Inches(0.25)

    def _point_list(self, slide, points: List[Point], x, y, w, h, max_pt=22, min_pt=12,
                    numbered=False, start=1):
        """Bulleted (or numbered) list of bold heading + explanation."""
        if not points:
            return
        heads = [p.heading.rstrip(":.") for p in points]
        make = lambda ds: [(f"00  {hd}: {d}", 1.0) for hd, d in zip(heads, ds)]  # noqa: E731
        details = shrink_to_fit([p.detail for p in points], make, w, h, min_pt, self.t.body_font)
        size = fit_size(make(details), w, h, max_pt, min_pt, self.t.body_font)
        tf = self._frame(slide, x, y, w, h)
        for i, pt in enumerate(points):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.space_after = Pt(size * 0.55)
            p.line_spacing = 1.12
            if numbered:
                self._run(p, f"{start + i:02d}  ", size, self.accent, self.t.title_font, bold=True)
            else:
                self._run(p, "■  ", size * 0.7, self.accent, self.t.body_font)
            self._run(p, f"{heads[i]}: ", size, self.title_c, self.t.body_font, bold=True)
            self._run(p, details[i], size, self.text_c, self.t.body_font)

    def _point_size(self, points: List[Point], w, h, max_pt, min_pt) -> int:
        items = [(f"00  {p.heading}: {p.detail}", 1.0) for p in points]
        return fit_size(items, w, h, max_pt, min_pt, self.t.body_font)

    def _point_columns(self, slide, points: List[Point], x, y, w, h, max_pt=22, min_pt=12,
                       numbered=False):
        """Two side-by-side lists that share ONE font size, split to balance text length."""
        gap = Inches(0.6)
        cw = (w - gap) / 2
        lens = [len(p.heading) + len(p.detail) for p in points]
        split = min(range(1, len(points)),
                    key=lambda k: (abs(sum(lens[:k]) - sum(lens[k:])), -k))
        cols = [points[:split], points[split:]]
        size = min(self._point_size(c, cw, h, max_pt, min_pt) for c in cols)
        self._point_list(slide, cols[0], x, y, cw, h, size, min(min_pt, size), numbered)
        self._point_list(slide, cols[1], x + cw + gap, y, cw, h, size, min(min_pt, size),
                         numbered, start=split + 1)

    def _string_list(self, slide, items: List[str], x, y, w, h, color, max_pt=22, min_pt=12):
        make = lambda ts: [(f"00 {t}", 1.0) for t in ts]  # noqa: E731
        items = shrink_to_fit(items, make, w, h, min_pt, self.t.body_font)
        size = fit_size(make(items), w, h, max_pt, min_pt, self.t.body_font)
        tf = self._frame(slide, x, y, w, h)
        for i, s in enumerate(items):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.space_after = Pt(size * 0.7)
            p.line_spacing = 1.1
            self._run(p, "●  ", size * 0.75, color, self.t.body_font)
            self._run(p, s, size, self.text_c, self.t.body_font)

    # ============================================================ slide types
    def cover(self, deck: Deck):
        s = self._new_slide()
        photo = fetch_image(deck.cover_image_query)
        if photo:
            self._picture_fill(s, photo[0], W * 0.55, 0, W * 0.45, H)
            self._rect(s, W * 0.55 - Inches(0.12), 0, Inches(0.12), H, self.accent)
            text_w = W * 0.55 - MARGIN - Inches(0.5)
            self._notes(s, None, photo[1])
        else:  # decorative circles instead of a photo
            d = H * 1.25
            circle = s.shapes.add_shape(MSO_SHAPE.OVAL, int(W * 0.62), int(-H * 0.12), int(d), int(d))
            circle.fill.solid()
            circle.fill.fore_color.rgb = rgb(mix(self.bg, self.accent, 0.14))
            circle.line.fill.background()
            small = s.shapes.add_shape(MSO_SHAPE.OVAL, int(W * 0.56), int(H * 0.62), int(H * 0.3), int(H * 0.3))
            small.fill.solid()
            small.fill.fore_color.rgb = rgb(mix(self.bg, self.accent2, 0.3))
            small.line.fill.background()
            text_w = W * 0.58 - MARGIN
        font = self.t.title_font
        size = fit_size([(deck.title, 1.0)], text_w, Inches(2.2), 48, 30, font)
        self._rect(s, MARGIN, Inches(2.1), Inches(1.5), Inches(0.1), self.accent)
        self._text(s, deck.title, MARGIN, Inches(2.35), text_w, Inches(2.2), size,
                   self.title_c, font, bold=True, anchor=MSO_ANCHOR.BOTTOM)
        sub = fit_size([(deck.subtitle, 1.0)], text_w, Inches(1.3), 22, 15, self.t.body_font)
        self._text(s, deck.subtitle, MARGIN, Inches(4.7), text_w, Inches(1.3), sub, self.muted)

    def section(self, spec: SlideSpec, number: int, section_no: int):
        s = self._new_slide()
        photo = fetch_image(spec.image_query)
        self._rect(s, 0, 0, Inches(0.25), H, self.accent)
        text_w = (W * 0.52 if photo else W - 2 * MARGIN) - MARGIN
        if photo:
            self._picture_fill(s, photo[0], W * 0.55, 0, W * 0.45, H)
        self._text(s, f"{section_no:02d}", MARGIN, Inches(1.6), text_w, Inches(1.4), 80,
                   mix(self.bg, self.accent, 0.55), self.t.title_font, bold=True)
        size = fit_size([(spec.title, 1.0)], text_w, Inches(1.5), 44, 28, self.t.title_font)
        self._text(s, spec.title, MARGIN, Inches(3.0), text_w, Inches(1.5), size,
                   self.title_c, self.t.title_font, bold=True)
        if spec.subtitle:
            self._text(s, spec.subtitle, MARGIN, Inches(4.6), text_w, Inches(1.2), 18, self.muted)
        self._footer(s, number)
        self._notes(s, spec.speaker_notes, photo[1] if photo else None)

    def bullets(self, spec: SlideSpec, number: int):
        s = self._new_slide()
        photo = fetch_image(spec.image_query) if self._fits_beside_photo(spec) else None
        if photo:
            img_x = W * PHOTO_X
            self._picture_fill(s, photo[0], img_x, 0, W - img_x, H)
            right = img_x - Inches(0.45)
        else:
            right = W - MARGIN
        width = right - MARGIN
        top = self._header(s, spec, width)
        bottom = self._callout(s, spec.callout, MARGIN, width) if spec.callout else CONTENT_BOTTOM
        pts = spec.points[:6]
        if photo or len(pts) <= 3:
            self._point_list(s, pts, MARGIN, top, width, bottom - top)
        else:  # two numbered columns use the full width without long lines
            self._point_columns(s, pts, MARGIN, top, width, bottom - top, numbered=True)
        self._footer(s, number)
        self._notes(s, spec.speaker_notes, photo[1] if photo else None)

    def _fits_beside_photo(self, spec: SlideSpec, min_readable: int = 14) -> bool:
        """Content beats decoration: only show a photo if the text stays readable next to it."""
        if not spec.image_query:
            return False
        width = W * PHOTO_X - Inches(0.45) - MARGIN
        height = CONTENT_BOTTOM - Inches(2.6) - (Inches(1.1) if spec.callout else 0)
        items = [(f"00  {p.heading}: {p.detail}", 1.0) for p in spec.points[:6]]
        return fits(items, width, height, min_readable, self.t.body_font)

    def cards(self, spec: SlideSpec, number: int):
        s = self._new_slide()
        top = self._header(s, spec, W - 2 * MARGIN)
        bottom = self._callout(s, spec.callout, MARGIN, W - 2 * MARGIN) if spec.callout else CONTENT_BOTTOM
        pts = spec.points[:6]
        cols = {1: 1, 2: 2, 3: 3, 4: 2}.get(len(pts), 3)   # 4 cards -> 2x2 grid
        rows = math.ceil(len(pts) / cols)
        cw = (W - 2 * MARGIN - GAP * (cols - 1)) / cols
        ch = (bottom - top - GAP * (rows - 1)) / rows
        pad = Inches(0.25)
        for i, pt in enumerate(pts):
            x = MARGIN + (i % cols) * (cw + GAP)
            y = top + (i // cols) * (ch + GAP)
            self._rect(s, x, y, cw, ch, self.surface, rounded=True, radius=0.05)
            self._rect(s, x, y, cw, Inches(0.07), self.accent if i % 2 == 0 else self.accent2)
            badge = Inches(0.5)
            wide = ch < Inches(2.6)  # short, wide card: number sits left of the text
            b = s.shapes.add_shape(MSO_SHAPE.OVAL, int(x + pad), int(y + pad + Inches(0.05)),
                                   int(badge), int(badge))
            b.fill.solid()
            b.fill.fore_color.rgb = rgb(self.accent)
            b.line.fill.background()
            b.shadow.inherit = False
            btf = b.text_frame
            btf.margin_left = btf.margin_right = btf.margin_top = btf.margin_bottom = 0
            bp = btf.paragraphs[0]
            bp.alignment = PP_ALIGN.CENTER
            self._run(bp, str(i + 1), 16, on_color(self.accent), self.t.title_font, bold=True)

            if wide:
                tx, ty = x + pad + badge + Inches(0.2), y + pad
            else:
                tx, ty = x + pad, y + pad + badge + Inches(0.15)
            tw, th = x + cw - pad - tx, y + ch - ty - pad
            hscale = 1.25 * 1.12 * max(1.0, CHAR_WIDTH.get(self.t.title_font, DEFAULT_CHAR_WIDTH)
                                       / CHAR_WIDTH.get(self.t.body_font, DEFAULT_CHAR_WIDTH))
            make = lambda ds: [(pt.heading, hscale), (ds[0], 1.0)]  # noqa: E731
            detail = shrink_to_fit([pt.detail], make, tw, th, 11, self.t.body_font)[0]
            size = fit_size(make([detail]), tw, th, 20, 11, self.t.body_font, gap=0.5)
            tf = self._frame(s, tx, ty, tw, th)
            p = tf.paragraphs[0]
            p.space_after = Pt(size * 0.5)
            self._run(p, pt.heading, round(size * 1.25), self.title_c, self.t.title_font, bold=True)
            p2 = tf.add_paragraph()
            p2.line_spacing = 1.08
            self._run(p2, detail, size, self.text_c, self.t.body_font)
        self._footer(s, number)
        self._notes(s, spec.speaker_notes)

    def two_column(self, spec: SlideSpec, number: int):
        s = self._new_slide()
        top = self._header(s, spec, W - 2 * MARGIN)
        bottom = self._callout(s, spec.callout, MARGIN, W - 2 * MARGIN) if spec.callout else CONTENT_BOTTOM
        col_w = (W - 2 * MARGIN - Inches(0.4)) / 2
        head_h = Inches(0.65)
        hsize = min(fit_size([(c.heading, 1.05)], col_w - Inches(0.5), head_h, 20, 13,
                             self.t.title_font, line=1.1) for c in (spec.left, spec.right))
        for i, (col, color) in enumerate(((spec.left, self.accent), (spec.right, self.accent2))):
            x = MARGIN + i * (col_w + Inches(0.4))
            self._rect(s, x, top, col_w, head_h, color)
            self._text(s, shorten(col.heading, 60), x + Inches(0.25), top, col_w - Inches(0.5),
                       head_h, hsize, on_color(color), self.t.title_font, bold=True,
                       anchor=MSO_ANCHOR.MIDDLE)
            self._rect(s, x, top + head_h, col_w, bottom - top - head_h, self.surface)
            self._string_list(s, col.points[:6], x + Inches(0.25), top + head_h + Inches(0.2),
                              col_w - Inches(0.5), bottom - top - head_h - Inches(0.35), color)
        self._footer(s, number)
        self._notes(s, spec.speaker_notes)

    def stats(self, spec: SlideSpec, number: int):
        s = self._new_slide()
        top = self._header(s, spec, W - 2 * MARGIN)
        source_y = CONTENT_BOTTOM - Inches(0.3)
        bottom = source_y - Inches(0.1)
        if spec.callout:
            bottom = self._callout(s, spec.callout, MARGIN, W - 2 * MARGIN, bottom)
        tiles = spec.stats[:4]
        n = len(tiles)
        tw = (W - 2 * MARGIN - GAP * (n - 1)) / n
        if spec.points:
            th = max(Inches(2.2), min(Inches(2.8), (bottom - top) * 0.5))
            ty = top
        else:  # no supporting points: bigger tiles, centred in the free space
            th = min(bottom - top, Inches(3.4))
            ty = top + (bottom - top - th) / 2
        inner = max(Inches(0.15), (th - Inches(2.5)) / 2)   # centre value+label inside tile
        label_h = th - inner - Inches(1.15)
        for i, st in enumerate(tiles):
            x = MARGIN + i * (tw + GAP)
            self._rect(s, x, ty, tw, th, self.surface, rounded=True, radius=0.06)
            self._rect(s, x, ty, Inches(0.08), th, self.palette[i % 2])
            vsize = fit_size([(st.value, 1.0)], tw - Inches(0.4), Inches(1.0), 60, 26,
                             self.t.title_font, line=1.1)
            self._text(s, st.value, x + Inches(0.25), ty + inner, tw - Inches(0.4),
                       Inches(1.0), vsize, self.palette[i % 2], self.t.title_font, bold=True,
                       anchor=MSO_ANCHOR.BOTTOM)
            make = lambda ts: [(ts[0], 1.0)]  # noqa: E731
            label = shrink_to_fit([st.label], make, tw - Inches(0.4), label_h, 11,
                                  self.t.body_font, floor=20)[0]
            lsize = fit_size(make([label]), tw - Inches(0.4), label_h, 20, 11, self.t.body_font)
            self._text(s, label, x + Inches(0.25), ty + inner + Inches(1.05), tw - Inches(0.4),
                       label_h, lsize, self.text_c)
        if spec.points:
            py = ty + th + Inches(0.35)
            pts = spec.points[:4]
            if len(pts) <= 2:
                self._point_list(s, pts, MARGIN, py, W - 2 * MARGIN, bottom - py, max_pt=20)
            else:
                self._point_columns(s, pts, MARGIN, py, W - 2 * MARGIN, bottom - py, max_pt=20)
        if spec.source:
            self._text(s, f"Source: {spec.source}", MARGIN, source_y, W - 2 * MARGIN,
                       Inches(0.3), 10, self.muted, italic=True)
        self._footer(s, number)
        self._notes(s, spec.speaker_notes)

    def chart(self, spec: SlideSpec, number: int):
        c: ChartSpec = spec.chart
        s = self._new_slide()
        top = self._header(s, spec, W - 2 * MARGIN)
        source_y = CONTENT_BOTTOM - Inches(0.3)
        bottom = source_y - Inches(0.1)
        insights = spec.points[:3]
        chart_w = (W - 2 * MARGIN) * (0.6 if insights else 1.0)

        data = CategoryChartData()
        data.categories = c.categories
        series_list = c.series[:1] if c.chart_type == "pie" else c.series
        for ser in series_list:
            data.add_series(ser.name, ser.values)
        frame = s.shapes.add_chart(CHART_TYPES[c.chart_type], int(MARGIN), int(top),
                                   int(chart_w), int(bottom - top), data)
        self._style_chart(frame.chart, c, len(series_list))

        if insights:
            px = MARGIN + chart_w + Inches(0.4)
            pw = W - MARGIN - px
            self._rect(s, px, top, pw, bottom - top, self.surface, rounded=True, radius=0.04)
            self._text(s, "KEY INSIGHTS", px + Inches(0.3), top + Inches(0.2), pw - Inches(0.6),
                       Inches(0.4), 13, self.accent, self.t.title_font, bold=True)
            self._point_list(s, insights, px + Inches(0.25), top + Inches(0.7), pw - Inches(0.5),
                             bottom - top - Inches(0.9), max_pt=19, min_pt=11)
        self._text(s, f"Source: {c.source}", MARGIN, source_y, W - 2 * MARGIN, Inches(0.3),
                   10, self.muted, italic=True)
        self._footer(s, number)
        self._notes(s, spec.speaker_notes)

    def _style_chart(self, ch, c: ChartSpec, n_series: int):
        ch.font.size, ch.font.name = Pt(13), self.t.body_font
        ch.font.color.rgb = rgb(self.text_c)
        ch.has_title = False
        ch.has_legend = c.chart_type == "pie" or n_series > 1
        if ch.has_legend:
            ch.legend.position = XL_LEGEND_POSITION.BOTTOM
            ch.legend.include_in_layout = False
        plot = ch.plots[0]
        if c.chart_type == "pie":
            for i, point in enumerate(plot.series[0].points):
                point.format.fill.solid()
                point.format.fill.fore_color.rgb = rgb(self.palette[i % len(self.palette)])
            plot.has_data_labels = True
            plot.data_labels.font.color.rgb = rgb(on_color(self.accent))
            plot.data_labels.font.bold = True
            return
        for i, ser in enumerate(plot.series):
            color = rgb(self.palette[i % len(self.palette)])
            if c.chart_type == "line":
                ser.format.line.color.rgb = color
                ser.format.line.width = Pt(3)
                ser.smooth = False
            else:
                ser.format.fill.solid()
                ser.format.fill.fore_color.rgb = color
        if c.chart_type in ("column", "bar"):
            plot.gap_width = 60
            if n_series == 1:
                plot.has_data_labels = True
                plot.data_labels.font.size = Pt(11)
                plot.data_labels.font.color.rgb = rgb(self.text_c)
        ch.value_axis.major_gridlines.format.line.color.rgb = rgb(mix(self.bg, self.text_c, 0.15))
        ch.value_axis.format.line.fill.background()
        ch.category_axis.format.line.color.rgb = rgb(mix(self.bg, self.text_c, 0.3))

    # ============================================================ entry point
    def build(self, deck: Deck, log: Callable[[str], None] = print) -> bytes:
        """Build the deck and return the .pptx file contents as bytes."""
        self.deck_title = deck.title
        log("Building cover slide")
        self.cover(deck)
        section_no = 0
        for i, spec in enumerate(deck.slides, 1):
            log(f"Building slide {i}/{len(deck.slides)}: {spec.title}")
            number = i + 1
            kind = _usable_kind(spec)
            if kind == "section":
                section_no += 1
                self.section(spec, number, section_no)
            else:
                getattr(self, kind)(spec, number)
        buf = io.BytesIO()
        self.prs.save(buf)
        return buf.getvalue()


def _chart_ok(c: Optional[ChartSpec]) -> bool:
    return bool(c and c.categories and c.series
                and all(len(s.values) == len(c.categories) for s in c.series))


def _usable_kind(spec: SlideSpec) -> str:
    """Fall back to a bullets slide when the data for a fancier layout is missing.
    Converts the available data into points so nothing the AI wrote is lost."""
    kind = spec.kind
    if kind == "chart" and not _chart_ok(spec.chart):
        kind = "bullets"
    elif kind == "cards" and len(spec.points) < 2:
        kind = "bullets"
    elif kind == "stats" and len(spec.stats or []) < 2:
        spec.points = [Point(heading=st.value, detail=st.label) for st in spec.stats or []] + spec.points
        kind = "bullets"
    elif kind == "two_column" and not (spec.left and spec.right and spec.left.points and spec.right.points):
        for col in (spec.left, spec.right):
            if col and col.points:
                spec.points.append(Point(heading=col.heading, detail="; ".join(col.points)))
        kind = "bullets"
    return kind
