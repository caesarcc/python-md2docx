import os
from pathlib import Path
from typing import List, Optional
import markdown
from pygments.formatters import HtmlFormatter
from xhtml2pdf import pisa

SUPPORTED_EXTENSIONS = (".md", ".markdown")


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
    """Build the destination PDF path.

    If overwrite is False, mirrors Windows Explorer behaviour:
    name.pdf, then "name (2).pdf", "name (3).pdf" and so on if the file already exists.
    """
    target_dir = output_dir if output_dir else source.parent
    target_dir.mkdir(parents=True, exist_ok=True)
    base_name = source.stem

    candidate = target_dir / f"{base_name}.pdf"
    if overwrite or not candidate.exists():
        return candidate

    counter = 2
    while True:
        candidate = target_dir / f"{base_name} ({counter}).pdf"
        if not candidate.exists():
            return candidate
        counter += 1


def build_css(
    font_size: int = 10,
    orientation: str = "vertical",
    page_size: str = "a4",
    custom_css: Optional[str] = None,
) -> str:
    """Generate the full CSS stylesheet for the PDF document."""
    page_orientation = "landscape" if orientation.lower() in ("horizontal", "landscape") else "portrait"
    pygments_css = HtmlFormatter(style="friendly").get_style_defs(".codehilite")

    base_css = f"""
    @page {{
        size: {page_size} {page_orientation};
        margin: 20mm 15mm 20mm 15mm;
        @frame footer {{
            -pdf-frame-content: footerContent;
            bottom: 8mm;
            margin-left: 15mm;
            margin-right: 15mm;
            height: 10mm;
        }}
    }}

    body {{
        font-family: Helvetica, Arial, sans-serif;
        font-size: {font_size}pt;
        line-height: 1.5;
        color: #24292f;
    }}

    h1, h2, h3, h4, h5, h6 {{
        color: #0f172a;
        font-weight: bold;
        margin-top: 1.2em;
        margin-bottom: 0.5em;
    }}

    h1 {{
        font-size: {font_size + 8}pt;
        border-bottom: 1px solid #d0d7de;
        padding-bottom: 4px;
    }}

    h2 {{
        font-size: {font_size + 4}pt;
        border-bottom: 1px solid #eaecef;
        padding-bottom: 3px;
    }}

    h3 {{
        font-size: {font_size + 2}pt;
    }}

    h4, h5, h6 {{
        font-size: {font_size}pt;
    }}

    p {{
        margin-top: 0;
        margin-bottom: 0.8em;
    }}

    a {{
        color: #0969da;
        text-decoration: underline;
    }}

    ul, ol {{
        margin-top: 0;
        margin-bottom: 0.8em;
        padding-left: 20px;
    }}

    li {{
        margin-bottom: 0.3em;
    }}

    blockquote {{
        margin: 1em 0;
        padding: 6px 12px;
        border-left: 4px solid #0969da;
        background-color: #f6f8fa;
        color: #57606a;
    }}

    table {{
        width: 100%;
        border-collapse: collapse;
        margin-top: 1em;
        margin-bottom: 1em;
    }}

    th, td {{
        border: 1px solid #d0d7de;
        padding: 6px 10px;
        text-align: left;
    }}

    th {{
        background-color: #f6f8fa;
        font-weight: bold;
    }}

    tr:nth-child(even) {{
        background-color: #fbfcfd;
    }}

    code {{
        font-family: Courier, monospace;
        font-size: {max(font_size - 1, 7)}pt;
        background-color: #f6f8fa;
        padding: 1px 3px;
        border-radius: 3px;
        color: #24292f;
    }}

    pre {{
        background-color: #f6f8fa;
        border: 1px solid #d0d7de;
        border-radius: 4px;
        padding: 8px 10px;
        font-family: Courier, monospace;
        font-size: {max(font_size - 1.5, 6.5)}pt;
        line-height: 1.4;
        margin-top: 0.8em;
        margin-bottom: 0.8em;
    }}

    pre code {{
        background-color: transparent;
        padding: 0;
        border: none;
    }}

    img {{
        max-width: 100%;
        height: auto;
    }}

    hr {{
        border: 0;
        border-top: 1px solid #d0d7de;
        margin: 1.5em 0;
    }}

    .codehilite {{
        background-color: #f6f8fa;
        border: 1px solid #d0d7de;
        border-radius: 4px;
        padding: 8px 10px;
        margin-top: 0.8em;
        margin-bottom: 0.8em;
    }}

    {pygments_css}
    """

    if custom_css:
        base_css += f"\n\n/* Custom User Styles */\n{custom_css}"

    return base_css


def make_link_callback(base_dir: Path):
    """Create a callback for xhtml2pdf to resolve relative image and asset paths."""
    def link_callback(uri: str, rel: str = "") -> str:
        # If absolute file or web URL, keep as is
        if uri.startswith(("http://", "https://", "data:")):
            return uri

        # Resolve relative to base_dir (the markdown file folder)
        resolved_path = (base_dir / uri).resolve()
        if resolved_path.exists():
            return str(resolved_path)

        # Fallback to current working directory
        cwd_path = (Path.cwd() / uri).resolve()
        if cwd_path.exists():
            return str(cwd_path)

        return uri

    return link_callback


def build_pdf(
    source: Path,
    dest: Path,
    orientation: str = "vertical",
    font_size: int = 10,
    page_size: str = "a4",
    css_path: Optional[Path] = None,
    overwrite: bool = False,
) -> Path:
    """Read a Markdown file and render it to a PDF document at dest."""
    with open(source, "r", encoding="utf-8-sig") as f:
        md_text = f.read()

    html_body = markdown.markdown(
        md_text,
        extensions=[
            "extra",
            "codehilite",
            "tables",
            "toc",
            "sane_lists",
            "nl2br",
        ],
    )

    custom_css = None
    if css_path:
        with open(css_path, "r", encoding="utf-8") as f:
            custom_css = f.read()

    stylesheet = build_css(
        font_size=font_size,
        orientation=orientation,
        page_size=page_size,
        custom_css=custom_css,
    )

    document_html = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<title>{source.stem}</title>
<style>
{stylesheet}
</style>
</head>
<body>
{html_body}
</body>
</html>"""

    dest_path = dest if overwrite else resolve_output_path(source, dest.parent, overwrite=False)

    link_callback = make_link_callback(source.parent)

    with open(dest_path, "wb") as output_file:
        pisa_status = pisa.CreatePDF(
            document_html,
            dest=output_file,
            link_callback=link_callback,
            encoding="utf-8",
        )

    if pisa_status.err:
        raise RuntimeError(f"Erro ao gerar PDF para {source.name} (código de erro: {pisa_status.err})")

    return dest_path
