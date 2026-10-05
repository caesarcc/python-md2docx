"""Core conversion logic for converting Markdown files to Microsoft Word (.docx) documents.
Supports rich markdown elements, table styling, code blocks, math symbols, and SVG diagram embedding.
"""

import base64
import io
import os
import re
import sys
import urllib.request
import zlib
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Tuple, Union

import docx
from docx import Document
from docx.enum.section import WD_ORIENTATION
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.opc.constants import RELATIONSHIP_TYPE
from docx.opc.packuri import PackURI
from docx.opc.part import Part
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.shared import Inches, Pt, RGBColor
from PIL import Image
import markdown
from bs4 import BeautifulSoup, NavigableString, Tag

SUPPORTED_EXTENSIONS = (".md", ".markdown")

PAGE_SIZES = {
    "a4": (Inches(8.27), Inches(11.69)),
    "letter": (Inches(8.5), Inches(11.0)),
    "legal": (Inches(8.5), Inches(14.0)),
}

MATH_REPLACEMENTS = {
    r"\times": "×",
    r"\to": "→",
    r"\cap": "∩",
    r"\cup": "∪",
    r"\emptyset": "∅",
    r"\in": "∈",
    r"\notin": "∉",
    r"\le": "≤",
    r"\leq": "≤",
    r"\ge": "≥",
    r"\geq": "≥",
    r"\ne": "≠",
    r"\neq": "≠",
    r"\approx": "≈",
    r"\pm": "±",
    r"\cdot": "·",
    r"\infty": "∞",
    r"\subset": "⊂",
    r"\subseteq": "⊆",
    r"\forall": "∀",
    r"\exists": "∃",
    r"\alpha": "α",
    r"\beta": "β",
    r"\gamma": "γ",
    r"\delta": "δ",
    r"\pi": "π",
    r"\sigma": "σ",
    r"\theta": "θ",
    r"\lambda": "λ",
}


def find_input_files(input_path: Path, pattern: str = "*.md") -> List[Path]:
    """Resolve the list of Markdown files to convert.

    If input_path is a file, it is returned as-is (regardless of pattern).
    If input_path is a directory, it is searched using the glob pattern,
    keeping only files with a supported Markdown extension.
    """
    if input_path.is_file():
        return [input_path]
    if not input_path.is_dir():
        raise FileNotFoundError(f"Caminho não encontrado: {input_path}")

    matches = sorted(input_path.glob(pattern))
    return [f for f in matches if f.is_file() and f.suffix.lower() in SUPPORTED_EXTENSIONS]


def resolve_output_path(source: Path, output_dir: Optional[Path] = None, overwrite: bool = False) -> Path:
    """Build the destination DOCX path.

    If overwrite is False, mirrors Windows Explorer behaviour:
    name.docx, then "name (2).docx", "name (3).docx" and so on if the file already exists.
    """
    target_dir = output_dir if output_dir else source.parent
    target_dir.mkdir(parents=True, exist_ok=True)
    base_name = source.stem

    candidate = target_dir / f"{base_name}.docx"
    if overwrite or not candidate.exists():
        return candidate

    counter = 2
    while True:
        candidate = target_dir / f"{base_name} ({counter}).docx"
        if not candidate.exists():
            return candidate
        counter += 1


def clean_inline_math(text: str) -> str:
    """Clean common LaTeX expressions to clean readable Unicode symbols."""
    def replace_math(match):
        expr = match.group(1)
        for cmd, repl in MATH_REPLACEMENTS.items():
            expr = expr.replace(cmd, repl)
        expr = re.sub(r"\\(?:text|mathrm|mathbf)\{([^}]+)\}", r"\1", expr)
        expr = re.sub(r"\\([a-zA-Z]+)", r"\1", expr)
        return expr.strip()

    return re.sub(r"\$([^\$]+)\$", replace_math, text)


def parse_css_overrides(css_text: Optional[str]) -> Dict[str, str]:
    """Extract basic styling overrides (colors, fonts) from a CSS stylesheet."""
    overrides = {}
    if not css_text:
        return overrides

    rules = re.findall(r"([^{]+)\{([^}]+)\}", css_text)
    for selector, body in rules:
        selector = selector.strip().lower()
        decls = {}
        for decl in body.split(";"):
            if ":" in decl:
                prop, val = decl.split(":", 1)
                decls[prop.strip().lower()] = val.strip()

        if "body" in selector:
            if "font-family" in decls:
                fonts = [f.strip().strip("'\"") for f in decls["font-family"].split(",")]
                overrides["font_family"] = fonts[0]
            if "color" in decls:
                c = decls["color"].lstrip("#")
                if len(c) == 6:
                    overrides["body_color"] = c.upper()
        if "h1" in selector and "color" in decls:
            c = decls["color"].lstrip("#")
            if len(c) == 6:
                overrides["h1_color"] = c.upper()
        if "h2" in selector and "color" in decls:
            c = decls["color"].lstrip("#")
            if len(c) == 6:
                overrides["h2_color"] = c.upper()
        if "blockquote" in selector:
            border_c = decls.get("border-left-color") or decls.get("border-color")
            if border_c:
                c = border_c.lstrip("#")
                if len(c) == 6:
                    overrides["quote_border"] = c.upper()

    return overrides


def render_diagram_to_svg_and_raster(code: str, diagram_type: str = "mermaid") -> Tuple[Optional[bytes], Optional[bytes]]:
    """Renders a diagram (mermaid, plantuml, dot, etc.) to SVG bytes and a raster fallback."""
    svg_bytes = None
    raster_bytes = None
    cleaned_code = code.strip()

    # Strategy 1: mermaid.ink (for mermaid)
    if diagram_type.lower() == "mermaid":
        try:
            encoded = base64.urlsafe_b64encode(cleaned_code.encode("utf-8")).decode("ascii")
            url_svg = f"https://mermaid.ink/svg/{encoded}"
            req = urllib.request.Request(url_svg, headers={"User-Agent": "Mozilla/5.0"})
            with urllib.request.urlopen(req, timeout=12) as resp:
                if resp.status == 200:
                    svg_bytes = resp.read()
        except Exception:
            pass

        try:
            url_img = f"https://mermaid.ink/img/{encoded}"
            req = urllib.request.Request(url_img, headers={"User-Agent": "Mozilla/5.0"})
            with urllib.request.urlopen(req, timeout=12) as resp:
                if resp.status == 200:
                    im = Image.open(io.BytesIO(resp.read()))
                    png_buf = io.BytesIO()
                    im.save(png_buf, format="PNG")
                    raster_bytes = png_buf.getvalue()
        except Exception:
            pass

    # Strategy 2: Kroki.io (supports mermaid, plantuml, graphviz, etc.)
    if not svg_bytes:
        try:
            compressed = zlib.compress(cleaned_code.encode("utf-8"), 9)
            kroki_encoded = base64.urlsafe_b64encode(compressed).decode("ascii")
            kroki_type = diagram_type.lower()
            if kroki_type == "dot":
                kroki_type = "graphviz"
            url_kroki = f"https://kroki.io/{kroki_type}/svg/{kroki_encoded}"
            req = urllib.request.Request(url_kroki, headers={"User-Agent": "Mozilla/5.0"})
            with urllib.request.urlopen(req, timeout=12) as resp:
                if resp.status == 200:
                    svg_bytes = resp.read()
        except Exception:
            pass

        if not raster_bytes:
            try:
                url_kroki_png = f"https://kroki.io/{kroki_type}/png/{kroki_encoded}"
                req = urllib.request.Request(url_kroki_png, headers={"User-Agent": "Mozilla/5.0"})
                with urllib.request.urlopen(req, timeout=12) as resp:
                    if resp.status == 200:
                        raster_bytes = resp.read()
            except Exception:
                pass

    # Fallback blank raster image if we only have svg_bytes
    if svg_bytes and not raster_bytes:
        img = Image.new("RGBA", (1000, 600), (255, 255, 255, 0))
        buf = io.BytesIO()
        img.save(buf, format="PNG")
        raster_bytes = buf.getvalue()

    return svg_bytes, raster_bytes


def add_svg_picture_to_paragraph(paragraph, svg_bytes: bytes, fallback_png_bytes: bytes, width=None, height=None):
    """Embeds an SVG image into a Word document paragraph using DrawingML svgBlip extension."""
    doc_part = paragraph.part
    package = doc_part.package

    png_stream = io.BytesIO(fallback_png_bytes)
    run = paragraph.add_run()
    inline = run.add_picture(png_stream, width=width, height=height)

    media_parts = [p for p in package.parts if p.partname.startswith("/word/media/")]
    next_idx = len(media_parts) + 1
    svg_partname = PackURI(f"/word/media/image{next_idx}.svg")
    svg_part = Part(svg_partname, "image/svg+xml", svg_bytes, package=package)
    package.parts.append(svg_part)
    svg_r_id = doc_part.relate_to(svg_part, RELATIONSHIP_TYPE.IMAGE)

    blip = inline._inline.xpath(".//a:blip")[0]
    ext_lsts = blip.xpath("./a:extLst")
    if ext_lsts:
        ext_lst = ext_lsts[0]
    else:
        ext_lst = parse_xml('<a:extLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>')
        blip.append(ext_lst)

    svg_ext = parse_xml(
        f'<a:ext xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        f'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        f'xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" '
        f'uri="{{96DAC542-70C3-44E3-80E4-348542EBD27B}}">'
        f'<asvg:svgBlip r:embed="{svg_r_id}"/>'
        f'</a:ext>'
    )
    ext_lst.append(svg_ext)
    return inline


def add_hyperlink(paragraph, url: str, text: str, color="0969DA", underline=True, bold=False, italic=False):
    """Adds a clickable hyperlink with custom styling to a paragraph."""
    part = paragraph.part
    r_id = part.relate_to(url, RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
    hyperlink = parse_xml(
        f'<w:hyperlink {nsdecls("w")} r:id="{r_id}" '
        f'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>'
    )
    new_run = parse_xml(f'<w:r {nsdecls("w")}/>')
    rPr = parse_xml(f'<w:rPr {nsdecls("w")}/>')
    if color:
        c = parse_xml(f'<w:color {nsdecls("w")} w:val="{color}"/>')
        rPr.append(c)
    if underline:
        u = parse_xml(f'<w:u {nsdecls("w")} w:val="single"/>')
        rPr.append(u)
    if bold:
        b = parse_xml(f'<w:b {nsdecls("w")}/>')
        rPr.append(b)
    if italic:
        i = parse_xml(f'<w:i {nsdecls("w")}/>')
        rPr.append(i)
    new_run.append(rPr)
    r_text = parse_xml(f'<w:t {nsdecls("w")}/>')
    r_text.text = text
    new_run.append(r_text)
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)


def set_run_code_style(run, font_name="Consolas", font_size_pt=9, fill_hex="F1F5F9", color_hex="1E293B"):
    """Applies inline code styling with monospace font and subtle background shading."""
    run.font.name = font_name
    run.font.size = Pt(font_size_pt)
    run.font.color.rgb = RGBColor.from_string(color_hex)
    rPr = run._r.get_or_add_rPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:val="clear" w:color="auto" w:fill="{fill_hex}"/>')
    rPr.append(shd)


def set_heading_bottom_border(paragraph, color="D0D7DE", sz="6"):
    """Adds a GitHub-style subtle horizontal accent border below headings."""
    pPr = paragraph._p.get_or_add_pPr()
    pBdr = parse_xml(
        f'<w:pBdr {nsdecls("w")}>'
        f'<w:bottom w:val="single" w:sz="{sz}" w:space="4" w:color="{color}"/>'
        f'</w:pBdr>'
    )
    pPr.append(pBdr)


def add_horizontal_rule(doc, color="D0D7DE", sz="6"):
    """Adds a clean horizontal divider line."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(8)
    pPr = p._p.get_or_add_pPr()
    pBdr = parse_xml(
        f'<w:pBdr {nsdecls("w")}>'
        f'<w:bottom w:val="single" w:sz="{sz}" w:space="1" w:color="{color}"/>'
        f'</w:pBdr>'
    )
    pPr.append(pBdr)


class MarkdownToDocxConverter:
    """Translates Markdown syntax tree into a Microsoft Word (.docx) document."""

    def __init__(
        self,
        orientation: str = "vertical",
        font_size: int = 10,
        page_size: str = "a4",
        output_dir: Optional[Path] = None,
        source_path: Optional[Path] = None,
        css_overrides: Optional[Dict[str, str]] = None,
    ):
        self.orientation = orientation.lower()
        self.font_size = font_size
        self.page_size = page_size.lower()
        self.source_path = source_path
        self.source_dir = source_path.parent if source_path else Path.cwd()
        self.output_dir = output_dir if output_dir else self.source_dir
        self.css_overrides = css_overrides or {}
        self.diagram_counter = 0

        self.doc = Document()
        self._setup_page()
        self._setup_styles()
        self._setup_footer()

    def _setup_page(self):
        section = self.doc.sections[0]
        # Margins
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.9)
        section.right_margin = Inches(0.9)

        base_w, base_h = PAGE_SIZES.get(self.page_size, PAGE_SIZES["a4"])
        if self.orientation in ("horizontal", "landscape"):
            section.orientation = WD_ORIENTATION.LANDSCAPE
            section.page_width = max(base_w, base_h)
            section.page_height = min(base_w, base_h)
        else:
            section.orientation = WD_ORIENTATION.PORTRAIT
            section.page_width = min(base_w, base_h)
            section.page_height = max(base_w, base_h)

        self.printable_width = section.page_width - section.left_margin - section.right_margin

    def _setup_styles(self):
        normal = self.doc.styles["Normal"]
        font_family = self.css_overrides.get("font_family", "Calibri")
        normal.font.name = font_family
        normal.font.size = Pt(self.font_size)

        body_hex = self.css_overrides.get("body_color", "24292F")
        normal.font.color.rgb = RGBColor.from_string(body_hex)
        normal.paragraph_format.line_spacing = 1.15
        normal.paragraph_format.space_after = Pt(4)
        normal.paragraph_format.space_before = Pt(0)

    def _setup_footer(self):
        """Configure page numbers in the document footer."""
        section = self.doc.sections[0]
        footer = section.footer
        p = footer.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)

        run = p.add_run("Página ")
        run.font.name = self.css_overrides.get("font_family", "Calibri")
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(0x94, 0xA3, 0xB8)

        fld_page = parse_xml(f'<w:fldSimple {nsdecls("w")} w:instr="PAGE"/>')
        p._p.append(fld_page)

    def convert(self, md_content: str) -> Document:
        """Parse markdown string and render elements into Document."""
        cleaned_md = clean_inline_math(md_content)

        html = markdown.markdown(
            cleaned_md,
            extensions=[
                "extra",
                "tables",
                "fenced_code",
                "sane_lists",
                "nl2br",
                "toc",
            ],
        )

        soup = BeautifulSoup(html, "html.parser")
        for element in soup.children:
            if isinstance(element, Tag):
                self._render_element(element)

        return self.doc

    def _render_element(self, tag: Tag):
        name = tag.name.lower()

        if name in ("h1", "h2", "h3", "h4", "h5", "h6"):
            self._render_heading(tag)
        elif name == "p":
            self._render_paragraph(tag)
        elif name == "hr":
            add_horizontal_rule(self.doc)
        elif name in ("ul", "ol"):
            self._render_list(tag, is_ordered=(name == "ol"), level=0)
        elif name == "blockquote":
            self._render_blockquote(tag)
        elif name == "table":
            self._render_table(tag)
        elif name == "pre":
            self._render_code_or_diagram(tag)
        else:
            for child in tag.children:
                if isinstance(child, Tag):
                    self._render_element(child)

    def _render_heading(self, tag: Tag):
        level = int(tag.name[1])
        p = self.doc.add_paragraph()
        p.paragraph_format.keep_with_next = True

        h1_color = self.css_overrides.get("h1_color", "0F172A")
        h2_color = self.css_overrides.get("h2_color", "1E293B")

        sizes = {
            1: (self.font_size + 8, h1_color, 18, 6, "D0D7DE", "6"),
            2: (self.font_size + 4, h2_color, 14, 4, "EAECEF", "4"),
            3: (self.font_size + 2, "334155", 10, 3, None, None),
            4: (self.font_size + 1, "475569", 8, 2, None, None),
            5: (self.font_size, "475569", 6, 2, None, None),
            6: (max(self.font_size - 1, 8), "64748B", 6, 2, None, None),
        }
        fsize, color_hex, before, after, border_color, border_sz = sizes.get(level, (self.font_size, "0F172A", 6, 2, None, None))

        p.paragraph_format.space_before = Pt(before)
        p.paragraph_format.space_after = Pt(after)

        self._render_inline(tag, p, is_bold=True, color_hex=color_hex, size_pt=fsize)

        if border_color:
            set_heading_bottom_border(p, color=border_color, sz=border_sz)

    def _render_paragraph(self, tag: Tag):
        imgs = tag.find_all("img", recursive=False)
        if len(imgs) == 1 and len(tag.get_text(strip=True)) == 0:
            self._render_image(imgs[0])
            return

        p = self.doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(4)
        self._render_inline(tag, p)

    def _render_image(self, img_tag: Tag):
        src = img_tag.get("src", "")
        if not src:
            return

        p = self.doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(6)
        p.paragraph_format.space_after = Pt(6)

        try:
            if src.startswith(("http://", "https://")):
                req = urllib.request.Request(src, headers={"User-Agent": "Mozilla/5.0"})
                img_data = urllib.request.urlopen(req, timeout=10).read()
                img_stream = io.BytesIO(img_data)
            else:
                img_path = (self.source_dir / src).resolve()
                if not img_path.exists():
                    img_path = (Path.cwd() / src).resolve()
                if not img_path.exists():
                    p.add_run(f"[Imagem não encontrada: {src}]")
                    return
                img_stream = str(img_path)

            p.add_run().add_picture(img_stream, width=min(self.printable_width, Inches(6)))
        except Exception as e:
            p.add_run(f"[Erro ao carregar imagem {src}: {e}]")

    def _render_list(self, list_tag: Tag, is_ordered: bool = False, level: int = 0):
        for li in list_tag.find_all("li", recursive=False):
            nested_lists = li.find_all(["ul", "ol"], recursive=False)

            p = self.doc.add_paragraph()
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(2)

            indent_inches = 0.25 * (level + 1)
            p.paragraph_format.left_indent = Inches(indent_inches)

            if is_ordered:
                style_name = "List Number" if level == 0 else f"List Number {min(level + 1, 3)}"
            else:
                style_name = "List Bullet" if level == 0 else f"List Bullet {min(level + 1, 3)}"

            try:
                p.style = self.doc.styles[style_name]
            except KeyError:
                bullet_char = "• " if not is_ordered else f"{level + 1}. "
                p.add_run(bullet_char)

            # Check task lists
            text = li.get_text()
            if text.startswith("[ ] "):
                p.add_run("☐ ").bold = True
            elif text.startswith(("[x] ", "[X] ")):
                p.add_run("☑ ").bold = True

            for child in li.children:
                if child in nested_lists:
                    continue
                if isinstance(child, NavigableString):
                    raw = str(child)
                    if raw.startswith(("[ ] ", "[x] ", "[X] ")):
                        raw = raw[4:]
                    p.add_run(raw)
                elif isinstance(child, Tag):
                    self._render_inline(child, p)

            for sublist in nested_lists:
                self._render_list(sublist, is_ordered=(sublist.name == "ol"), level=level + 1)

    def _render_blockquote(self, tag: Tag):
        table = self.doc.add_table(rows=1, cols=1)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        cell = table.cell(0, 0)
        cell.width = self.printable_width

        quote_border_color = self.css_overrides.get("quote_border", "0969DA")

        tcPr = cell._tc.get_or_add_tcPr()
        tcPr.append(parse_xml(f'<w:shd {nsdecls("w")} w:val="clear" w:color="auto" w:fill="F8FAFC"/>'))
        borders = parse_xml(
            f'<w:tcBorders {nsdecls("w")}>'
            f'<w:top w:val="none"/>'
            f'<w:left w:val="single" w:sz="24" w:space="0" w:color="{quote_border_color}"/>'
            f'<w:bottom w:val="none"/>'
            f'<w:right w:val="none"/>'
            f'</w:tcBorders>'
        )
        tcPr.append(borders)
        tcMar = parse_xml(
            f'<w:tcMar {nsdecls("w")}>'
            f'<w:top w:w="120" w:type="dxa"/>'
            f'<w:bottom w:w="120" w:type="dxa"/>'
            f'<w:left w:w="200" w:type="dxa"/>'
            f'<w:right w:w="160" w:type="dxa"/>'
            f'</w:tcMar>'
        )
        tcPr.append(tcMar)

        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(2)
        self._render_inline(tag, p, is_italic=True, color_hex="475569")

        sp = self.doc.add_paragraph()
        sp.paragraph_format.space_before = Pt(0)
        sp.paragraph_format.space_after = Pt(4)

    def _render_code_or_diagram(self, pre_tag: Tag):
        code_tag = pre_tag.find("code")
        raw_code = code_tag.get_text() if code_tag else pre_tag.get_text()

        classes = []
        if code_tag and code_tag.get("class"):
            classes.extend(code_tag.get("class"))
        if pre_tag.get("class"):
            classes.extend(pre_tag.get("class"))

        lang = ""
        for cls in classes:
            if cls.startswith("language-"):
                lang = cls.replace("language-", "").lower()
                break

        if lang in ("mermaid", "plantuml", "dot", "graphviz"):
            self._render_diagram(raw_code, lang)
        else:
            self._render_code_block(raw_code, lang)

    def _render_diagram(self, code: str, lang: str):
        self.diagram_counter += 1
        base_name = self.source_path.stem if self.source_path else "diagram"
        svg_filename = f"{base_name}_diagram_{self.diagram_counter}.svg"
        svg_file_path = self.output_dir / svg_filename

        svg_bytes, raster_bytes = render_diagram_to_svg_and_raster(code, lang)

        if svg_bytes:
            try:
                self.output_dir.mkdir(parents=True, exist_ok=True)
                svg_file_path.write_bytes(svg_bytes)
                print(f"Diagrama SVG gerado: {svg_file_path.name}")
            except Exception as e:
                print(f"Erro ao salvar SVG {svg_file_path}: {e}", file=sys.stderr)

            p = self.doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(8)
            p.paragraph_format.space_after = Pt(8)

            try:
                diag_width = min(self.printable_width, Inches(6.0))
                add_svg_picture_to_paragraph(p, svg_bytes, raster_bytes, width=diag_width)
                return
            except Exception as exc:
                print(f"Erro ao incorporar SVG no docx: {exc}", file=sys.stderr)

        self._render_code_block(f"[Diagrama: {lang}]\n{code}", lang)

    def _render_code_block(self, code: str, lang: str = ""):
        table = self.doc.add_table(rows=1, cols=1)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        cell = table.cell(0, 0)
        cell.width = self.printable_width

        tcPr = cell._tc.get_or_add_tcPr()
        tcPr.append(parse_xml(f'<w:shd {nsdecls("w")} w:val="clear" w:color="auto" w:fill="F6F8FA"/>'))
        borders = parse_xml(
            f'<w:tcBorders {nsdecls("w")}>'
            f'<w:top w:val="single" w:sz="4" w:space="0" w:color="D0D7DE"/>'
            f'<w:left w:val="single" w:sz="4" w:space="0" w:color="D0D7DE"/>'
            f'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="D0D7DE"/>'
            f'<w:right w:val="single" w:sz="4" w:space="0" w:color="D0D7DE"/>'
            f'</w:tcBorders>'
        )
        tcPr.append(borders)
        tcMar = parse_xml(
            f'<w:tcMar {nsdecls("w")}>'
            f'<w:top w:w="120" w:type="dxa"/>'
            f'<w:bottom w:w="120" w:type="dxa"/>'
            f'<w:left w:w="160" w:type="dxa"/>'
            f'<w:right w:w="160" w:type="dxa"/>'
            f'</w:tcMar>'
        )
        tcPr.append(tcMar)

        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.15

        lines = code.strip("\n").split("\n")
        for i, line in enumerate(lines):
            run = p.add_run(line)
            run.font.name = "Consolas"
            run.font.size = Pt(max(self.font_size - 1.5, 8.0))
            run.font.color.rgb = RGBColor(0x1E, 0x29, 0x3B)
            if i < len(lines) - 1:
                p.add_run("\n")

        sp = self.doc.add_paragraph()
        sp.paragraph_format.space_before = Pt(0)
        sp.paragraph_format.space_after = Pt(4)

    def _render_table(self, table_tag: Tag):
        rows = table_tag.find_all("tr")
        if not rows:
            return

        num_rows = len(rows)
        num_cols = max(len(r.find_all(["th", "td"])) for r in rows)

        table = self.doc.add_table(rows=num_rows, cols=num_cols)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER

        tblPr = table._tbl.tblPr
        tblPr.append(parse_xml(
            f'<w:tblBorders {nsdecls("w")}>'
            f'<w:top w:val="single" w:sz="4" w:space="0" w:color="D0D7DE"/>'
            f'<w:left w:val="single" w:sz="4" w:space="0" w:color="D0D7DE"/>'
            f'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="D0D7DE"/>'
            f'<w:right w:val="single" w:sz="4" w:space="0" w:color="D0D7DE"/>'
            f'<w:insideH w:val="single" w:sz="4" w:space="0" w:color="D0D7DE"/>'
            f'<w:insideV w:val="single" w:sz="4" w:space="0" w:color="D0D7DE"/>'
            f'</w:tblBorders>'
        ))

        alignments = []
        first_row_cells = rows[0].find_all(["th", "td"])
        for c in first_row_cells:
            align_attr = c.get("align", "").lower()
            style_attr = c.get("style", "").lower()
            if "center" in align_attr or "text-align: center" in style_attr:
                alignments.append(WD_ALIGN_PARAGRAPH.CENTER)
            elif "right" in align_attr or "text-align: right" in style_attr:
                alignments.append(WD_ALIGN_PARAGRAPH.RIGHT)
            else:
                alignments.append(WD_ALIGN_PARAGRAPH.LEFT)

        col_width = self.printable_width / num_cols

        for r_idx, r_tag in enumerate(rows):
            row = table.rows[r_idx]
            trPr = row._tr.get_or_add_trPr()
            trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))

            is_header = bool(r_tag.find("th")) or (r_idx == 0)
            if is_header:
                trPr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))

            fill_hex = "F6F8FA" if is_header else ("F8FAFC" if r_idx % 2 == 1 else "FFFFFF")

            cells = r_tag.find_all(["th", "td"])
            for c_idx, c_tag in enumerate(cells):
                if c_idx >= num_cols:
                    break
                cell = row.cells[c_idx]
                cell.width = col_width

                tcPr = cell._tc.get_or_add_tcPr()
                tcPr.append(parse_xml(f'<w:shd {nsdecls("w")} w:val="clear" w:color="auto" w:fill="{fill_hex}"/>'))
                tcPr.append(parse_xml(
                    f'<w:tcMar {nsdecls("w")}>'
                    f'<w:top w:w="120" w:type="dxa"/>'
                    f'<w:bottom w:w="120" w:type="dxa"/>'
                    f'<w:left w:w="160" w:type="dxa"/>'
                    f'<w:right w:w="160" w:type="dxa"/>'
                    f'</w:tcMar>'
                ))
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

                p = cell.paragraphs[0]
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(0)
                if c_idx < len(alignments):
                    p.alignment = alignments[c_idx]

                self._render_inline(c_tag, p, is_bold=is_header)

        sp = self.doc.add_paragraph()
        sp.paragraph_format.space_before = Pt(0)
        sp.paragraph_format.space_after = Pt(4)

    def _render_inline(
        self,
        node: Union[Tag, NavigableString],
        paragraph,
        is_bold: bool = False,
        is_italic: bool = False,
        is_strike: bool = False,
        is_underline: bool = False,
        is_code: bool = False,
        color_hex: Optional[str] = None,
        size_pt: Optional[float] = None,
        link_url: Optional[str] = None,
    ):
        for child in node.children:
            if isinstance(child, NavigableString):
                text = str(child)
                if not text:
                    continue

                if link_url:
                    add_hyperlink(
                        paragraph,
                        link_url,
                        text,
                        color="0969DA",
                        underline=True,
                        bold=is_bold,
                        italic=is_italic,
                    )
                else:
                    run = paragraph.add_run(text)
                    if is_bold:
                        run.bold = True
                    if is_italic:
                        run.italic = True
                    if is_strike:
                        run.font.strike = True
                    if is_underline:
                        run.underline = True
                    if color_hex:
                        run.font.color.rgb = RGBColor.from_string(color_hex)
                    if size_pt:
                        run.font.size = Pt(size_pt)
                    if is_code:
                        set_run_code_style(run, font_size_pt=max(self.font_size - 1, 8.5))

            elif isinstance(child, Tag):
                cname = child.name.lower()
                next_bold = is_bold or (cname in ("strong", "b"))
                next_italic = is_italic or (cname in ("em", "i"))
                next_strike = is_strike or (cname in ("del", "s"))
                next_underline = is_underline or (cname == "u")
                next_code = is_code or (cname == "code")
                next_link = child.get("href") if cname == "a" else link_url

                if cname == "br":
                    paragraph.add_run().add_break()
                elif cname == "img":
                    self._render_image(child)
                else:
                    self._render_inline(
                        child,
                        paragraph,
                        is_bold=next_bold,
                        is_italic=next_italic,
                        is_strike=next_strike,
                        is_underline=next_underline,
                        is_code=next_code,
                        color_hex=color_hex,
                        size_pt=size_pt,
                        link_url=next_link,
                    )


def build_docx(
    source: Path,
    dest: Path,
    orientation: str = "vertical",
    font_size: int = 10,
    page_size: str = "a4",
    css_path: Optional[Path] = None,
    overwrite: bool = False,
) -> Path:
    """Read a Markdown file and render it to a DOCX document at dest."""
    with open(source, "r", encoding="utf-8-sig") as f:
        md_text = f.read()

    css_overrides = {}
    if css_path and css_path.is_file():
        try:
            with open(css_path, "r", encoding="utf-8") as f:
                css_overrides = parse_css_overrides(f.read())
        except Exception as e:
            print(f"Aviso: Não foi possível ler o arquivo CSS ({e})", file=sys.stderr)

    dest_path = dest if overwrite else resolve_output_path(source, dest.parent, overwrite=False)

    converter = MarkdownToDocxConverter(
        orientation=orientation,
        font_size=font_size,
        page_size=page_size,
        output_dir=dest_path.parent,
        source_path=source,
        css_overrides=css_overrides,
    )

    doc = converter.convert(md_text)
    doc.save(str(dest_path))
    return dest_path


# Compatibility alias
build_pdf = build_docx
