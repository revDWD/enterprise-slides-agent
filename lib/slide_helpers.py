"""Helpers for building template-based, hand-drawn-style decks with python-pptx.

Everything here produces native, editable PowerPoint objects (no bitmaps).
"""
import copy
import random
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"


# ---------- template handling ----------

def open_template(template_path, out_path):
    """Open a COPY of the template. Refuses to write over the template itself."""
    template_path, out_path = Path(template_path).resolve(), Path(out_path).resolve()
    if template_path == out_path:
        raise ValueError("Output path must differ from the template path.")
    prs = Presentation(str(template_path))
    return prs


def describe_layouts(prs):
    """Return [(layout_name, [(placeholder_idx, type, name)])] to map content to layouts."""
    out = []
    for layout in prs.slide_layouts:
        phs = [(p.placeholder_format.idx, str(p.placeholder_format.type), p.name)
               for p in layout.placeholders]
        out.append((layout.name, phs))
    return out


def get_layout(prs, name):
    for layout in prs.slide_layouts:
        if layout.name.lower() == name.lower():
            return layout
    raise KeyError(f"Layout '{name}' not found. Available: {[l.name for l in prs.slide_layouts]}")


def delete_slide(prs, index):
    """Delete slide at index, dropping its relationship so no orphan part is saved."""
    sld_id_lst = prs.slides._sldIdLst
    sld_id = list(sld_id_lst)[index]
    prs.part.drop_rel(sld_id.rId)
    sld_id_lst.remove(sld_id)


def delete_all_slides(prs):
    for i in range(len(prs.slides) - 1, -1, -1):
        delete_slide(prs, i)


# ---------- accessibility ----------

def set_alt_text(shape, text, title=None):
    c_nv_pr = shape._element.xpath(".//p:cNvPr")[0]
    c_nv_pr.set("descr", text)
    if title:
        c_nv_pr.set("title", title)


def mark_decorative(shape):
    """Flag a shape as decorative (Office 2019+ `a16:decorative` extension)."""
    c_nv_pr = shape._element.xpath(".//p:cNvPr")[0]
    c_nv_pr.set("descr", "")
    ext_lst = c_nv_pr.find(qn("a:extLst"))
    if ext_lst is None:
        ext_lst = etree.SubElement(c_nv_pr, qn("a:extLst"))
    ext = etree.SubElement(ext_lst, qn("a:ext"), uri="{C183D7F6-B498-43B3-948B-1728B52AA6E4}")
    etree.SubElement(
        ext,
        "{http://schemas.microsoft.com/office/drawing/2017/decorative}decorative",
        val="1",
        nsmap={"adec": "http://schemas.microsoft.com/office/drawing/2017/decorative"},
    )


def set_notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text


# ---------- hand-drawn doodles (native freeform shapes with bezier curves) ----------

def _add_custgeom_shape(slide, left, top, width, height, path_cmds, color, weight_pt,
                        name, fill=None, closed=False, arrow_end=False):
    """Insert a freeform whose path is given as ('M',pt)/('C',p1,p2,p3)/('L',pt) in a
    0..1 normalised space scaled to the shape box."""
    shape = slide.shapes.add_shape(1, left, top, width, height)  # rectangle, replaced geometry
    shape.name = name
    sp_pr = shape._element.spPr
    prst = sp_pr.find(qn("a:prstGeom"))
    W, H = 100000, 100000

    def pt(p):
        return f'<a:pt x="{int(p[0] * W)}" y="{int(p[1] * H)}"/>'

    parts = []
    for cmd in path_cmds:
        if cmd[0] == "M":
            parts.append(f"<a:moveTo>{pt(cmd[1])}</a:moveTo>")
        elif cmd[0] == "L":
            parts.append(f"<a:lnTo>{pt(cmd[1])}</a:lnTo>")
        elif cmd[0] == "C":
            parts.append(f"<a:cubicBezTo>{pt(cmd[1])}{pt(cmd[2])}{pt(cmd[3])}</a:cubicBezTo>")
    if closed:
        parts.append("<a:close/>")
    xml = (
        f'<a:custGeom xmlns:a="{A_NS}"><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/>'
        f'<a:rect l="0" t="0" r="r" b="b"/><a:pathLst>'
        f'<a:path w="{W}" h="{H}"{"" if fill else " fill=\"none\""}>' + "".join(parts) +
        "</a:path></a:pathLst></a:custGeom>"
    )
    geom = etree.fromstring(xml)
    sp_pr.replace(prst, geom)

    if fill:
        shape.fill.solid()
        shape.fill.fore_color.rgb = RGBColor.from_string(fill)
    else:
        shape.fill.background()
    shape.line.color.rgb = RGBColor.from_string(color)
    shape.line.width = Pt(weight_pt)
    ln = shape.line._get_or_add_ln()
    ln.set("cap", "rnd")
    etree.SubElement(ln, qn("a:round"))
    if arrow_end:
        etree.SubElement(ln, qn("a:tailEnd"), type="triangle", w="med", len="med")
    # drop theme style effects (shadows) so the doodle stays flat
    style = shape._element.find(qn("p:style"))
    if style is not None:
        shape._element.remove(style)
    mark_decorative(shape)
    return shape


def squiggle_underline(slide, left, top, width, color="E65050", weight_pt=3.5,
                       waves=None, amplitude=0.28, seed=7, name="Doodle - squiggle underline"):
    """Wavy hand-drawn underline. `top` is the baseline; height is derived from width."""
    rnd = random.Random(seed)
    waves = waves or max(3, int(width / Emu(914400 * 0.45)))
    height = int(max(Emu(914400 * 0.12), width * 0.04))
    cmds = [("M", (0.0, 0.5 + rnd.uniform(-0.1, 0.1)))]
    step = 1.0 / (waves * 2)
    x = 0.0
    up = True
    for _ in range(waves * 2):
        x2 = min(1.0, x + step)
        jy = rnd.uniform(-0.08, 0.08)
        peak = 0.5 + (-amplitude if up else amplitude) * 1.6 + jy
        peak = max(0.0, min(1.0, peak))
        c1 = (x + step * 0.25 + rnd.uniform(-0.01, 0.01), peak)
        c2 = (x + step * 0.75 + rnd.uniform(-0.01, 0.01), peak)
        end = (x2, 0.5 + rnd.uniform(-0.08, 0.08))
        cmds.append(("C", c1, c2, end))
        x = x2
        up = not up
    return _add_custgeom_shape(slide, left, top, width, height, cmds, color, weight_pt, name)


def hand_circle(slide, left, top, width, height, color="143565", weight_pt=3.0, seed=3,
                name="Doodle - circle"):
    """Slightly overshooting hand-drawn ellipse to ring a key number."""
    rnd = random.Random(seed)
    j = lambda: rnd.uniform(-0.03, 0.03)
    cmds = [
        ("M", (0.50 + j(), 0.03 + j())),
        ("C", (0.85, 0.00), (1.00, 0.30), (0.97 + j(), 0.55)),
        ("C", (0.95, 0.85), (0.70, 1.00), (0.45 + j(), 0.97)),
        ("C", (0.15, 0.95), (0.00, 0.70), (0.03 + j(), 0.42)),
        ("C", (0.05, 0.12), (0.30, 0.00), (0.58 + j(), 0.06)),
        ("C", (0.68, 0.08), (0.75, 0.11), (0.80 + j(), 0.16)),
    ]
    return _add_custgeom_shape(slide, left, top, width, height, cmds, color, weight_pt, name)


def hand_arrow(slide, left, top, width, height, color="143565", weight_pt=3.0, seed=5,
               name="Doodle - arrow"):
    """Curved hand-drawn arrow from top-left to bottom-right of the box."""
    rnd = random.Random(seed)
    j = lambda: rnd.uniform(-0.04, 0.04)
    cmds = [
        ("M", (0.0, 0.1 + j())),
        ("C", (0.35, 0.0 + j()), (0.75, 0.2 + j()), (1.0, 0.9)),
    ]
    return _add_custgeom_shape(slide, left, top, width, height, cmds, color, weight_pt, name,
                               arrow_end=True)


def journey_path(slide, left, top, width, height, color="3EB3B5", weight_pt=4.0, seed=11,
                 name="Doodle - journey path"):
    """Meandering left-to-right squiggle path for roadmaps/timelines."""
    rnd = random.Random(seed)
    cmds = [("M", (0.0, 0.6))]
    xs = [0.2, 0.4, 0.6, 0.8, 1.0]
    for i, x in enumerate(xs):
        y = 0.2 + 0.6 * (i % 2) + rnd.uniform(-0.1, 0.1)
        cmds.append(("C", (x - 0.15, 0.5 + rnd.uniform(-0.4, 0.4)),
                     (x - 0.05, y + rnd.uniform(-0.3, 0.3)), (x, max(0.05, min(0.95, y)))))
    return _add_custgeom_shape(slide, left, top, width, height, cmds, color, weight_pt, name,
                               arrow_end=True)


# ---------- QA ----------

def audit(pptx_path, template_layout_names=None, min_font_pt=10):
    """Structural QA. Returns a list of issue strings (empty list == clean)."""
    prs = Presentation(str(pptx_path))
    issues = []
    allowed = set(template_layout_names) if template_layout_names else None
    for n, slide in enumerate(prs.slides, 1):
        if allowed and slide.slide_layout.name not in allowed:
            issues.append(f"slide {n}: layout '{slide.slide_layout.name}' not in template")
        if slide.shapes.title is None or not slide.shapes.title.has_text_frame \
                or not slide.shapes.title.text_frame.text.strip():
            issues.append(f"slide {n}: missing/empty title placeholder")
        for sh in slide.shapes:
            if sh.is_placeholder and sh.has_text_frame and not sh.text_frame.text.strip() \
                    and sh.placeholder_format.type is not None:
                issues.append(f"slide {n}: empty placeholder '{sh.name}'")
            if sh.shape_type == 13:  # picture
                cnv = sh._element.xpath(".//p:cNvPr")[0]
                if not cnv.get("descr"):
                    issues.append(f"slide {n}: picture '{sh.name}' lacks alt text")
            if sh.has_text_frame:
                for p in sh.text_frame.paragraphs:
                    for r in p.runs:
                        if r.font.size is not None and r.font.size.pt < min_font_pt:
                            issues.append(f"slide {n}: '{sh.name}' font {r.font.size.pt}pt < {min_font_pt}pt")
                        if "[TBD" in r.text:
                            issues.append(f"slide {n}: unresolved placeholder text in '{sh.name}'")
            left, top = sh.left or 0, sh.top or 0
            if left < 0 or top < 0 or left + (sh.width or 0) > prs.slide_width \
                    or top + (sh.height or 0) > prs.slide_height:
                issues.append(f"slide {n}: '{sh.name}' extends beyond slide bounds")
        if slide.has_notes_slide is False and n > 1:
            issues.append(f"slide {n}: no speaker notes")
    return issues
