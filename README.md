# md2docx

Converte arquivos Markdown (`.md` / `.markdown`) em documentos Word do Microsoft Office (`.docx`) com alta fidelidade visual, preservando a experiência de visualização do Markdown. Inclui suporte nativo para conversão automática de diagramas (Mermaid, PlantUML, Graphviz) em arquivos vetoriais **SVG** salvos no disco e incorporados diretamente no documento Word.

## Recursos e Fidelidade Visual

- **Compatibilidade Word (.docx):** Gera arquivos DOCX modernos compatíveis com Word 365, Word 2016-2021, Word Online e LibreOffice.
- **Diagramas Vetoriais SVG:** Quando um diagrama (ex: ````mermaid ... ````) está presente no Markdown:
  - Gera e salva um arquivo `.svg` de alta resolução no disco.
  - Incorpora o SVG no documento Word usando extensões nativas DrawingML (`asvg:svgBlip`) com fallback rasterizado para máxima compatibilidade.
- **Títulos e Estrutura:** Cabeçalhos `H1` a `H6` estilizados no padrão moderno com linhas divisórias de destaque (`H1`, `H2`) e prevenção contra linhas órfãs (`keep_with_next`).
- **Tabelas Estilizadas:** Cabeçalho com sombreamento diferenciado, repetição em múltiplas páginas (`tblHeader`), prevenção de quebra de linhas (`cantSplit`), alinhamento de colunas (`:---`, `:---:`, `---:`) e zebra striping.
- **Blocos de Código e Sintaxe:** Blocos de código em container sombreado com fonte monoespaçada (`Consolas`), bordas suaves e preservação de recuos e quebras.
- **Código Inline:** Destaque monoespaçado com fundo cinza suave (`#F1F5F9`).
- **Fórmulas e Expressões Matemáticas:** Normalização e renderização limpa de expressões LaTeX e símbolos matemáticos comuns (`∩`, `∅`, `×`, `→`, `≤`, `≥`, etc.).
- **Listas e Tarefas:** Listas com marcadores e numeradas com níveis de indentação preservados e caixas de seleção interativas (`☐`, `☑`).
- **Citações (`blockquote`):** Caixas de destaque com borda lateral colorida e fundo suave.
- **Rodapé com Paginação:** Numeração dinâmica de páginas no rodapé do documento.
- **Imagens:** Resolução e ajuste automático de imagens locais e URLs externas à largura imprimível da página.

## Instalação

Requer Python 3.9+.

```bash
# A partir da raiz do projeto
python -m venv .venv
.venv\Scripts\activate        # Windows (cmd/PowerShell)
# source .venv/bin/activate   # Linux/macOS

pip install -e .
```

Isso instala as dependências (`markdown`, `python-docx`, `beautifulsoup4`, `pillow`, `pygments`) e disponibiliza o comando `md2docx` (e o alias `md2pdf`) no terminal.

Alternativamente, para usar diretamente sem instalar o pacote:

```bash
pip install -r requirements.txt
python -m md2docx <argumentos>
```

## Uso Básico

```bash
md2docx [entrada] [opções]
```

- `entrada`: pasta ou arquivo Markdown específico. **Padrão: pasta atual (`.`)**.
  - Se for uma pasta, todos os arquivos `.md`/`.markdown` que casarem com `--pattern` serão convertidos.
  - Se for um arquivo específico, apenas ele é convertido.

### Exemplos

Converter todos os arquivos Markdown da pasta atual:

```bash
md2docx
```

Converter todos os arquivos de uma pasta específica:

```bash
md2docx C:\Documentos
```

Converter um único arquivo:

```bash
md2docx C:\Documentos\relatorio.md
```

Usar um filtro (wildcard) para selecionar apenas alguns arquivos da pasta:

```bash
md2docx C:\Documentos --pattern "relatorio_2026_*.md"
```

Gerar documento Word em orientação horizontal (paisagem):

```bash
md2docx C:\Documentos\relatorio.md --orientation horizontal
```

Definir o tamanho da fonte base (padrão: 10pt):

```bash
md2docx C:\Documentos\relatorio.md --font-size 11
```

Definir o formato da página (A4, Letter ou Legal):

```bash
md2docx C:\Documentos\relatorio.md --page-size letter
```

Aplicar um arquivo CSS com estilos personalizados (cores, fontes e bordas):

```bash
md2docx C:\Documentos\relatorio.md --css estilos.css
```

Salvar os arquivos Word e diagramas SVG em uma pasta de destino específica:

```bash
md2docx C:\Documentos --output-dir C:\Documentos\Word
```

Sobrescrever o documento caso já exista no destino:

```bash
md2docx C:\Documentos\relatorio.md --overwrite
```

## Opções da Linha de Comando

| Opção | Atalho | Padrão | Descrição |
|---|---|---|---|
| `entrada` | — | `.` (pasta atual) | Pasta ou arquivo Markdown de entrada. |
| `--pattern` | `-p` | `*.md` | Wildcard para filtrar arquivos quando `entrada` é uma pasta. |
| `--orientation` | `-o` | `vertical` | Orientação da página: `vertical` (retrato) ou `horizontal` (paisagem). |
| `--font-size` | `-f` | `10` | Tamanho base da fonte, em pontos, usado no corpo do documento. |
| `--page-size` | — | `a4` | Tamanho da página: `a4`, `letter` ou `legal`. |
| `--css` | `-c` | nenhum | Caminho para arquivo `.css` com estilos personalizados. |
| `--output-dir` | `-d` | pasta do arquivo de origem | Pasta onde os documentos `.docx` e diagramas `.svg` serão salvos. |
| `--overwrite` | `-w` | desativado | Sobrescreve o documento existente (padrão: cria `nome (2).docx`, `nome (3).docx`, etc.). |
| `--version` | `-v` | — | Exibe a versão do programa. |
| `--help` | `-h` | — | Exibe a mensagem de ajuda com todos os comandos e opções. |

## Resolução de Nomes de Saída

Por padrão, o arquivo DOCX gerado recebe o mesmo nome do arquivo de origem com extensão `.docx`. Se o arquivo já existir no destino, o programa não sobrescreve por padrão: ele cria réplicas seguras `nome (2).docx`, `nome (3).docx`, etc., idêntico ao comportamento do Windows Explorer.

Para sobrescrever diretamente o arquivo existente, passe a flag `-w` ou `--overwrite`.

## Estrutura do Projeto

```
md2docx/
├── __init__.py     # expõe main(), build_docx() e __version__
├── __main__.py     # ponto de entrada executável: python -m md2docx
├── cli.py          # tratamento de argumentos e execução CLI
└── core.py         # motor de conversão Markdown -> DOCX e geração SVG
tests/
└── test_md2docx.py # suíte completa de testes unitários com pytest
pyproject.toml      # metadados e configuração de empacotamento
requirements.txt    # lista de dependências do projeto
```
