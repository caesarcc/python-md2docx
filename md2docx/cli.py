import argparse
import sys
from pathlib import Path
from typing import Optional, Sequence

from . import __version__
from .core import build_docx, find_input_files, resolve_output_path


def parse_args(argv: Optional[Sequence[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        prog="md2docx",
        description=(
            "Converte arquivos Markdown (.md, .markdown) em documentos Word (.docx) formatados com suporte a diagramas SVG."
        ),
    )
    parser.add_argument(
        "input",
        nargs="?",
        default=".",
        help="Pasta ou arquivo Markdown de entrada (padrão: pasta atual).",
    )
    parser.add_argument(
        "-p",
        "--pattern",
        default="*.md",
        help="Padrão wildcard para filtrar arquivos quando 'input' é uma pasta (padrão: '*.md').",
    )
    parser.add_argument(
        "-o",
        "--orientation",
        choices=["vertical", "horizontal"],
        default="vertical",
        help="Orientação da página: vertical ou horizontal (padrão: vertical).",
    )
    parser.add_argument(
        "-f",
        "--font-size",
        type=int,
        default=10,
        help="Tamanho base da fonte em pontos (padrão: 10).",
    )
    parser.add_argument(
        "--page-size",
        choices=["a4", "letter", "legal"],
        default="a4",
        help="Tamanho da página: a4, letter ou legal (padrão: a4).",
    )
    parser.add_argument(
        "-c",
        "--css",
        default=None,
        help="Caminho para arquivo .css com estilos personalizados.",
    )
    parser.add_argument(
        "-d",
        "--output-dir",
        default=None,
        help="Pasta de destino para os DOCXs gerados (padrão: mesma pasta do arquivo de origem).",
    )
    parser.add_argument(
        "-w",
        "--overwrite",
        action="store_true",
        help="Sobrescrever o DOCX de destino caso já exista (padrão: cria 'nome (2).docx').",
    )
    parser.add_argument(
        "-v",
        "--version",
        action="version",
        version=f"%(prog)s {__version__}",
    )
    return parser.parse_args(argv)


def main(argv: Optional[Sequence[str]] = None) -> int:
    args = parse_args(argv)

    input_path = Path(args.input).expanduser().resolve()
    output_dir = Path(args.output_dir).expanduser().resolve() if args.output_dir else None
    css_path = Path(args.css).expanduser().resolve() if args.css else None

    if args.font_size <= 0:
        print("Erro: --font-size deve ser um número positivo.", file=sys.stderr)
        return 1

    if css_path and not css_path.is_file():
        print(f"Erro: Arquivo CSS não encontrado: {css_path}", file=sys.stderr)
        return 1

    try:
        files = find_input_files(input_path, args.pattern)
    except FileNotFoundError as exc:
        print(f"Erro: {exc}", file=sys.stderr)
        return 1

    if not files:
        print("Nenhum arquivo Markdown encontrado com os critérios informados.", file=sys.stderr)
        return 1

    exit_code = 0
    for source in files:
        dest = resolve_output_path(source, output_dir, overwrite=args.overwrite)
        try:
            actual_dest = build_docx(
                source=source,
                dest=dest,
                orientation=args.orientation,
                font_size=args.font_size,
                page_size=args.page_size,
                css_path=css_path,
                overwrite=args.overwrite,
            )
            print(f"OK: {source.name} -> {actual_dest.name}")
        except Exception as exc:
            print(f"Falha ao converter {source.name}: {exc}", file=sys.stderr)
            exit_code = 1

    return exit_code


if __name__ == "__main__":
    raise SystemExit(main())
