# md2pdf

Converte arquivos Markdown (`.md` / `.markdown`) em documentos PDF elegantes e formatados, com suporte a tabelas, imagens locais, blocos de código com destaque de sintaxe, listas, citações e estilos personalizados via CSS.

## Instalação

Requer Python 3.9+.

```bash
# A partir da raiz do projeto
python -m venv .venv
.venv\Scripts\activate        # Windows (cmd/PowerShell)
# source .venv/bin/activate   # Linux/macOS

pip install -e .
```

Isso instala as dependências (`markdown`, `xhtml2pdf`, `pygments`) e disponibiliza o comando `md2pdf` no terminal.

Alternativamente, para usar sem instalar o pacote:

```bash
pip install -r requirements.txt
python -m md2pdf <argumentos>
```

## Uso básico

```bash
md2pdf [entrada] [opções]
```

- `entrada`: pasta ou arquivo Markdown específico. **Padrão: pasta atual (`.`)**.
  - Se for uma pasta, todos os arquivos `.md`/`.markdown` que casarem com `--pattern` serão convertidos (um PDF por arquivo).
  - Se for um arquivo específico, apenas ele é convertido.

### Exemplos

Converter todos os arquivos Markdown da pasta atual:

```bash
md2pdf
```

Converter todos os arquivos de uma pasta específica:

```bash
md2pdf C:\Documentos
```

Converter um único arquivo:

```bash
md2pdf C:\Documentos\artigo.md
```

Usar um filtro (wildcard) para selecionar apenas alguns arquivos da pasta:

```bash
md2pdf C:\Documentos --pattern "relatorio_2024_*.md"
```

Gerar PDF em orientação horizontal (paisagem):

```bash
md2pdf C:\Documentos\artigo.md --orientation horizontal
```

Definir o tamanho da fonte base (padrão: 10pt):

```bash
md2pdf C:\Documentos\artigo.md --font-size 11
```

Definir o formato da página (A4, Letter ou Legal):

```bash
md2pdf C:\Documentos\artigo.md --page-size letter
```

Aplicar um arquivo CSS com estilos personalizados:

```bash
md2pdf C:\Documentos\artigo.md --css estilos.css
```

Salvar os PDFs em uma pasta diferente da origem:

```bash
md2pdf C:\Documentos --output-dir C:\Documentos\PDF
```

Sobrescrever o PDF caso já exista no destino:

```bash
md2pdf C:\Documentos\artigo.md --overwrite
```

## Opções

| Opção | Atalho | Padrão | Descrição |
|---|---|---|---|
| `entrada` | — | `.` (pasta atual) | Pasta ou arquivo Markdown de entrada. |
| `--pattern` | `-p` | `*.md` | Wildcard para filtrar arquivos quando `entrada` é uma pasta. Ignorado quando `entrada` é um arquivo específico. |
| `--orientation` | `-o` | `vertical` | Orientação da página do PDF: `vertical` (retrato) ou `horizontal` (paisagem). |
| `--font-size` | `-f` | `10` | Tamanho base da fonte, em pontos, usado no corpo do documento. |
| `--page-size` | — | `a4` | Tamanho da página: `a4`, `letter` ou `legal`. |
| `--css` | `-c` | nenhum | Caminho para arquivo `.css` com regras adicionais de estilo para o PDF. |
| `--output-dir` | `-d` | pasta do arquivo de origem | Pasta onde os PDFs gerados serão salvos. |
| `--overwrite` | `-w` | desativado | Sobrescreve o PDF existente se já existir no destino (em vez de gerar `nome (2).pdf`). |
| `--version` | `-v` | — | Exibe a versão do programa. |
| `--help` | `-h` | — | Exibe a mensagem de ajuda com todos os comandos e opções. |

## Recursos Suportados

- **Títulos e Estrutura:** Suporte a cabeçalhos `H1` a `H6` com separadores visuais elegantes.
- **Formatação de Texto:** Negrito, itálico, tachado, links e citações (`blockquote`).
- **Tabelas:** Tabelas estilizadas automaticamente com cabeçalho destacado e linhas alternadas.
- **Blocos de Código:** Realce de sintaxe colorido (syntax highlighting) para dezenas de linguagens via Pygments (ex: ````python ... ````).
- **Imagens e Mídias:** Resolução automática de caminhos de imagens relativas à pasta do arquivo `.md`.
- **Quebras de Página e Layout:** Margens automáticas de 20mm e controle de formato A4/Letter/Legal.

## Nome do Arquivo de Saída

Por padrão, o PDF gerado recebe o mesmo nome do arquivo de origem (com extensão `.pdf`). Se um arquivo com esse nome já existir no destino, o programa **não sobrescreve**: ele replica o comportamento do Windows Explorer criando `nome (2).pdf`, `nome (3).pdf`, etc.

Para sobrescrever diretamente o arquivo existente, use a opção `-w` ou `--overwrite`.

## Estrutura do Projeto

```
md2pdf/
├── __init__.py     # expõe main() e __version__
├── __main__.py      # permite `python -m md2pdf`
├── cli.py           # parsing de argumentos (argparse) e orquestração
└── core.py          # busca de arquivos, conversão HTML/CSS e geração do PDF
pyproject.toml        # empacotamento, metadados e dependências
requirements.txt       # dependências para uso direto sem instalar o pacote
```
