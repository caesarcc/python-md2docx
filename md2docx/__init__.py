__version__ = "0.2.0"

from .cli import main
from .core import build_docx, build_pdf, find_input_files, resolve_output_path

__all__ = ["main", "build_docx", "build_pdf", "find_input_files", "resolve_output_path", "__version__"]
