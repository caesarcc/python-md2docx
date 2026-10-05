---
name: md2docx
description: |
  Converte arquivos Markdown (.md, .markdown) em documentos Microsoft Word (.docx) elegantes e formatados.
  Use esta skill sempre que o usuário solicitar gerar, criar, exportar ou converter arquivos Markdown em documentos Word (.docx), relatórios em Word ou formato Office.
  Gera automaticamente diagramas (Mermaid, PlantUML, Graphviz) em arquivos vetoriais .svg e os embute no Word com alta fidelidade visual.
license: Apache-2.0
metadata:
  version: v1
  publisher: caesarcc
---

# Conversão de Markdown para Word (.docx) com md2docx

O `md2docx` é um utilitário de linha de comando instalado globalmente para converter arquivos Markdown em documentos Microsoft Word (`.docx`) modernos e com alta fidelidade visual, com suporte a diagramas vetoriais SVG, tabelas estilizadas, blocos de código, código inline e fórmulas matemáticas.

## Quando Utilizar Esta Skill

Utilize sempre que o usuário solicitar:
- "Gere / crie um documento Word baseado neste Markdown"
- "Converta este arquivo .md para .docx"
- "Exporte o relatório em formato Office / Word"
- "Gere um arquivo docx com os diagramas"

## Comandos Recomendados

Execute no terminal através de `run_command`:

```bash
# Converter um arquivo específico (recomendado usar -w para sobrescrever)
md2docx caminho/para/arquivo.md -w

# Converter com orientação paisagem (horizontal)
md2docx caminho/para/arquivo.md -o horizontal -w

# Especificar tamanho de fonte e página
md2docx caminho/para/arquivo.md -f 11 --page-size a4 -w

# Salvar em diretório específico
md2docx caminho/para/arquivo.md -d caminho/pasta_saida -w

# Converter todos os arquivos .md da pasta atual
md2docx . -w
```

## Recursos Automáticos
1. **Diagramas (Mermaid, PlantUML, Graphviz):** O `md2docx` gera arquivos `.svg` no disco (diretório de saída) e os embute com qualidade vetorial nativa DrawingML (`asvg:svgBlip`) e fallback PNG no `.docx`.
2. **Tabelas:** Cabeçalhos destacados (`#F1F5F9`), repetição entre páginas (`tblHeader`), sem quebra no meio da linha (`cantSplit`) e alinhamento de colunas.
3. **Blocos de Código:** Estilizados em container sombreado (`#F6F8FA`) com fonte monoespaçada `Consolas`.
4. **Fórmulas Matemáticas:** Converte expressões LaTeX comuns para caracteres Unicode legíveis (`∩`, `∅`, `×`, `→`, etc.).
5. **Citações e Listas:** Formatação de blockquote com barra lateral colorida e suporte a tarefas com checkbox (`☐`, `☑`).
