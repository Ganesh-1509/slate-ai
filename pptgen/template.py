"""Build the deck on top of a user's own PowerPoint template (.pptx / .potx).

The template's slide layouts (backgrounds, logos, fonts, colours, decorations) are
used for every slide. Titles go into the template's own title placeholders; the
content (points, cards, charts, stats...) is drawn in the template's colour scheme.
"""
import io
import zipfile
from typing import Dict, Optional

from lxml import etree
from pptx import Presentation
from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.util import Inches, Pt

from .builder import DeckBuilder, _hex, contrast, fit_size
from .models import Deck, SlideSpec, ThemeSpec

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
}
TITLE_TYPES = {PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE}
TEXT_TYPES = {PP_PLACEHOLDER.BODY, PP_PLACEHOLDER.SUBTITLE, PP_PLACEHOLDER.OBJECT}
IGNORED_TYPES = {PP_PLACEHOLDER.DATE, PP_PLACEHOLDER.FOOTER, PP_PLACEHOLDER.SLIDE_NUMBER}


class TemplateError(ValueError):
    """The uploaded file can't be used as a template."""


# ---------------------------------------------------------------- loading
def load_template(data: bytes) -> Presentation:
    """Open a .pptx or .potx file and remove its existing slides."""
    try:
        prs = Presentation(io.BytesIO(_potx_to_pptx(data)))
    except Exception as e:
        raise TemplateError("That file isn't a valid PowerPoint (.pptx / .potx) file.") from e
    if not len(prs.slide_layouts):
        raise TemplateError("That template has no slide layouts to build on.")
    sld_ids = prs.slides._sldIdLst
    for sld_id in list(sld_ids):
        prs.part.drop_rel(sld_id.rId)
        sld_ids.remove(sld_id)
    return prs


def _potx_to_pptx(data: bytes) -> bytes:
    """A .potx is a .pptx with a different content type - rewrite it so it opens."""
    template_type = b"presentationml.template.main+xml"
    try:
        zin = zipfile.ZipFile(io.BytesIO(data))
        types = zin.read("[Content_Types].xml")
    except Exception:
        return data  # let python-pptx report the problem
    if template_type not in types:
        return data
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            blob = zin.read(item.filename)
            if item.filename == "[Content_Types].xml":
                blob = blob.replace(template_type, b"presentationml.presentation.main+xml")
            zout.writestr(item, blob)
    return out.getvalue()


# ---------------------------------------------------------------- theme
def theme_from_template(prs: Presentation, fallback: ThemeSpec) -> ThemeSpec:
    """Read colours and fonts from the template's slide master theme."""
    master = prs.slide_master
    try:
        theme_xml = etree.fromstring(master.part.part_related_by(RT.THEME).blob)
    except Exception:
        return fallback

    scheme: Dict[str, str] = {}
    for el in theme_xml.findall(".//a:clrScheme/*", NS):
        name = etree.QName(el).localname
        srgb, sysc = el.find("a:srgbClr", NS), el.find("a:sysClr", NS)
        if srgb is not None:
            scheme[name] = "#" + srgb.get("val")
        elif sysc is not None and sysc.get("lastClr"):
            scheme[name] = "#" + sysc.get("lastClr")

    # The master's colour map says which scheme colours are background / text
    cmap_el = master._element.find("p:clrMap", NS)
    cmap = dict(cmap_el.attrib) if cmap_el is not None else {}

    def resolve(alias: str, default: str) -> str:
        return scheme.get(cmap.get(alias, alias), default)

    bg = _master_background(master, scheme, cmap) or resolve("bg1", fallback.background)
    text = resolve("tx1", fallback.text_color)
    title = resolve("tx2", text)
    if contrast(title, bg) < 3:
        title = text
    accent = scheme.get("accent1", fallback.accent)
    accent2 = scheme.get("accent2", fallback.accent2)
    if contrast(accent, bg) < 1.6:  # accent too close to the background
        accent, accent2 = accent2, scheme.get("accent3", fallback.accent2)

    def font(kind: str, default: str) -> str:
        el = theme_xml.find(f".//a:fontScheme/a:{kind}/a:latin", NS)
        face = el.get("typeface") if el is not None else ""
        return face if face and not face.startswith("+") else default

    # model_construct: template fonts may be outside our usual font list
    return ThemeSpec.model_construct(
        background="#" + _hex(bg), title_color="#" + _hex(title), text_color="#" + _hex(text),
        accent="#" + _hex(accent), accent2="#" + _hex(accent2),
        title_font=font("majorFont", fallback.title_font),
        body_font=font("minorFont", fallback.body_font),
    )


def _master_background(master, scheme, cmap) -> Optional[str]:
    bg = master._element.find(".//p:cSld/p:bg", NS)
    if bg is None:
        return None
    srgb = bg.find(".//a:srgbClr", NS)
    if srgb is not None:
        return "#" + srgb.get("val")
    sc = bg.find(".//a:schemeClr", NS)
    if sc is not None:
        name = sc.get("val")
        return scheme.get(cmap.get(name, name))
    return None  # picture / gradient background: fall back to the scheme colour


# ---------------------------------------------------------------- layouts
def _kinds(layout):
    return [ph.placeholder_format.type for ph in layout.placeholders
            if ph.placeholder_format.type not in IGNORED_TYPES]


def _pick_layouts(prs: Presentation):
    layouts = list(prs.slide_layouts)

    def has_title(lo):
        return any(k in TITLE_TYPES for k in _kinds(lo))

    def extras(lo):
        return [k for k in _kinds(lo) if k not in TITLE_TYPES]

    title_only = next((lo for lo in layouts if has_title(lo) and not extras(lo)), None)
    if title_only is None:  # closest thing: a title + the fewest other placeholders
        titled = [lo for lo in layouts if has_title(lo)]
        title_only = min(titled, key=lambda lo: len(extras(lo))) if titled else layouts[-1]

    cover = next((lo for lo in layouts if PP_PLACEHOLDER.CENTER_TITLE in _kinds(lo)), None) \
        or next((lo for lo in layouts if PP_PLACEHOLDER.SUBTITLE in _kinds(lo)), None) \
        or layouts[0]
    section = next((lo for lo in layouts if "section" in lo.name.lower() and has_title(lo)), None) \
        or cover
    return cover, title_only, section


# ---------------------------------------------------------------- builder
class TemplateDeckBuilder(DeckBuilder):
    def __init__(self, template: bytes, fallback_theme: ThemeSpec):
        prs = load_template(template)
        super().__init__(theme_from_template(prs, fallback_theme), prs=prs)
        self.cover_layout, self.content_layout, self.section_layout = _pick_layouts(prs)
        self.photo_full_bleed = False
        self.content_bottom = self._content_bottom()

    def _content_bottom(self):
        """Stop content above the template's footer area."""
        tops = [ph.top for ph in self.content_layout.placeholders
                if ph.placeholder_format.type in IGNORED_TYPES and ph.top and ph.top > self.H * 0.7]
        return (min(tops) - Inches(0.1)) if tops else self.H - Inches(0.75)

    # -------- slides from the template's layouts
    def _from_layout(self, layout, keep=TITLE_TYPES):
        slide = self.prs.slides.add_slide(layout)
        for ph in list(slide.placeholders):
            if ph.placeholder_format.type not in keep:
                ph._element.getparent().remove(ph._element)
        return slide

    def _new_slide(self):
        return self._from_layout(self.content_layout)

    @staticmethod
    def _title_ph(slide):
        return next((ph for ph in slide.placeholders if ph.placeholder_format.type in TITLE_TYPES), None)

    @staticmethod
    def _text_ph(slide):
        return next((ph for ph in slide.placeholders if ph.placeholder_format.type in TEXT_TYPES), None)

    def _header(self, slide, spec: SlideSpec, width) -> int:
        title = self._title_ph(slide)
        if title is None or title.top is None or title.top > self.H * 0.5:
            return super()._header(slide, spec, width)
        # Copy the inherited position first: changing only the width of a placeholder
        # that inherits its position would reset it to the top-left corner.
        left, top, w, h = title.left, title.top, title.width, title.height
        w = min(w, int(self.margin + width - left))  # leave room for a photo
        title.left, title.top, title.width, title.height = left, top, w, h
        title.text_frame.text = spec.title
        size = fit_size([(spec.title, 1.05)], w, h, 40, 22, self.t.title_font, line=1.0)
        for run in title.text_frame.paragraphs[0].runs:
            run.font.size = Pt(size)
        y = top + h + Inches(0.1)
        if spec.subtitle:
            sub = fit_size([(spec.subtitle, 1.0)], width, Inches(0.6), 18, 13, self.t.body_font)
            self._text(slide, spec.subtitle, self.margin, y, width, Inches(0.6), sub,
                       self.muted, italic=True)
            y += Inches(0.6)
        return y + Inches(0.15)

    def _footer(self, slide, number: int, light=False):
        pass  # the template's own footer / slide numbers apply

    def cover(self, deck: Deck):
        slide = self._from_layout(self.cover_layout, keep=TITLE_TYPES | TEXT_TYPES)
        self._fill(slide, deck.title, deck.subtitle)

    def section(self, spec: SlideSpec, number: int, section_no: int):
        slide = self._from_layout(self.section_layout, keep=TITLE_TYPES | TEXT_TYPES)
        self._fill(slide, spec.title, spec.subtitle or "")
        self._notes(slide, spec.speaker_notes)

    def _fill(self, slide, title_text: str, sub_text: str):
        title, body = self._title_ph(slide), self._text_ph(slide)
        if title is not None:
            title.text_frame.text = title_text
        else:
            self._text(slide, title_text, self.margin, self.H * 0.35, self.W - 2 * self.margin,
                       Inches(1.4), 40, self.title_c, self.t.title_font, bold=True)
        if body is not None:
            body.text_frame.text = sub_text
        elif sub_text:
            self._text(slide, sub_text, self.margin, self.H * 0.55, self.W - 2 * self.margin,
                       Inches(1.0), 18, self.muted)


def describe_template(data: bytes) -> str:
    """Short summary for the UI, e.g. '12 layouts, 16:9'. Raises TemplateError."""
    prs = load_template(data)
    ratio = prs.slide_width / prs.slide_height
    shape = "16:9" if abs(ratio - 16 / 9) < 0.05 else "4:3" if abs(ratio - 4 / 3) < 0.05 else "custom size"
    return f"{len(prs.slide_layouts)} layouts, {shape}"
