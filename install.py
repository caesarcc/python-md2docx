#!/usr/bin/env python3
"""Script de instalação global do md2docx para uso com uv e integração com o Antigravity / Gemini CLI."""

import os
import shutil
import subprocess
import sys
from pathlib import Path


if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")


def print_step(msg: str):
    print(f"\n==> {msg}")


def print_success(msg: str):
    print(f"  [OK] {msg}")


def print_warn(msg: str):
    print(f"  [AVISO] {msg}")


def print_error(msg: str):
    print(f"  [ERRO] {msg}")


def find_uv_executable() -> str:
    """Find uv executable in PATH or user local bin."""
    uv_cmd = shutil.which("uv")
    if uv_cmd:
        return uv_cmd

    # Check common user local bin on Windows
    user_local_bin_uv = Path.home() / ".local" / "bin" / "uv.exe"
    if user_local_bin_uv.is_file():
        return str(user_local_bin_uv)

    return ""


def install_tool_with_uv(repo_dir: Path, uv_cmd: str) -> bool:
    print_step("Instalando o pacote python-md2docx globalmente com uv...")

    cmd = [uv_cmd, "tool", "install", "--editable", str(repo_dir), "--force"]
    print(f"Executando: {' '.join(cmd)}")

    try:
        res = subprocess.run(cmd, check=True, capture_output=True, text=True)
        if res.stdout:
            print(res.stdout.strip())
        print_success("md2docx instalado com sucesso via 'uv tool'!")
        return True
    except subprocess.CalledProcessError as exc:
        print_error(f"Falha ao executar uv tool install: {exc}")
        if exc.stderr:
            print(exc.stderr)
        return False


def verify_cli_installation() -> bool:
    print_step("Verificando comando 'md2docx' no PATH...")
    exe_path = shutil.which("md2docx")
    if not exe_path:
        local_bin_exe = Path.home() / ".local" / "bin" / "md2docx.exe"
        if local_bin_exe.is_file():
            exe_path = str(local_bin_exe)

    if not exe_path:
        print_warn("O comando 'md2docx' ainda não foi detectado no PATH atual.")
        print_warn("Certifique-se de que '%USERPROFILE%\\.local\\bin' está nas Variáveis de Ambiente (PATH).")
        return False

    try:
        res = subprocess.run([exe_path, "--version"], capture_output=True, text=True, check=True)
        version_str = res.stdout.strip() or res.stderr.strip()
        print_success(f"Executável localizado em: {exe_path}")
        print_success(f"Versão confirmada: {version_str}")
        return True
    except Exception as exc:
        print_error(f"Erro ao testar executável md2docx: {exc}")
        return False


def configure_antigravity_ai(repo_dir: Path) -> bool:
    print_step("Configurando o Antigravity / Gemini CLI para reconhecer o md2docx...")

    global_config_dir = Path.home() / ".gemini" / "config"
    if not global_config_dir.exists():
        print_warn(f"Diretório global de configurações do Antigravity não encontrado: {global_config_dir}")
        print_warn("Criando diretório para instalar a skill globalmente...")
        global_config_dir.mkdir(parents=True, exist_ok=True)

    # 1. Configurar Skill global
    skill_source = repo_dir / "skills" / "md2docx" / "SKILL.md"
    target_skill_dir = global_config_dir / "skills" / "md2docx"
    target_skill_file = target_skill_dir / "SKILL.md"

    if skill_source.is_file():
        target_skill_dir.mkdir(parents=True, exist_ok=True)
        shutil.copy2(skill_source, target_skill_file)
        print_success(f"Skill do Antigravity instalada em:\n       {target_skill_file}")
    else:
        print_warn(f"Arquivo de skill de origem não encontrado em: {skill_source}")

    # 2. Configurar Regra global (rules)
    target_rules_dir = global_config_dir / "rules"
    target_rules_dir.mkdir(parents=True, exist_ok=True)
    target_rule_file = target_rules_dir / "md2docx.md"

    rule_content = """# Regra Global: Conversão de Markdown para Word (.docx)

Sempre que o usuário solicitar gerar, criar, exportar ou converter um documento Markdown (`.md`, `.markdown`) para Microsoft Word (`.docx`) ou formato Office Word:
1. **Utilize a ferramenta CLI `md2docx`** instalada na máquina (disponível globalmente).
2. **Comando padrão recomendado:**
   ```bash
   md2docx <caminho/arquivo.md> -w
   ```
3. **Opções úteis:**
   - Para orientação paisagem: `md2docx <arquivo.md> -o horizontal -w`
   - Para alterar tamanho de fonte e folha: `md2docx <arquivo.md> -f 11 --page-size a4 -w`
   - Para salvar em pasta específica: `md2docx <arquivo.md> -d <pasta_destino> -w`
4. **Diagramas SVG:** O `md2docx` gera arquivos `.svg` para quaisquer blocos de diagramas (Mermaid, PlantUML, etc.) e os incorpora com fidelidade vetorial diretamente no documento Word.
"""
    target_rule_file.write_text(rule_content, encoding="utf-8")
    print_success(f"Regra global do Antigravity instalada em:\n       {target_rule_file}")

    return True


def main():
    print("=" * 70)
    print(" Instalador Global do python-md2docx (com uv e Antigravity) ")
    print("=" * 70)

    repo_dir = Path(__file__).resolve().parent

    uv_cmd = find_uv_executable()
    if not uv_cmd:
        print_error("O executável 'uv' não foi encontrado no sistema!")
        print("Por favor, instale o uv ou verifique o PATH: https://docs.astral.sh/uv/")
        sys.exit(1)

    print_success(f"uv localizado em: {uv_cmd}")

    # 1. Instalar ferramenta com uv
    ok_tool = install_tool_with_uv(repo_dir, uv_cmd)
    if not ok_tool:
        print_error("A instalação do pacote falhou.")
        sys.exit(1)

    # 2. Verificar executável
    verify_cli_installation()

    # 3. Configurar IA local (Antigravity)
    configure_antigravity_ai(repo_dir)

    print("\n" + "=" * 70)
    print(" 🎉 INSTALAÇÃO CONCLUÍDA COM SUCESSO!")
    print("=" * 70)
    print("Agora você pode:")
    print("  1. Rodar 'md2docx arquivo.md -w' de qualquer pasta no terminal.")
    print("  2. Em qualquer projeto no Antigravity, simplesmente pedir à IA:")
    print("     'Gere um documento word baseado neste markdown'")
    print("     A IA local já possui a skill e a regra para usar o md2docx automaticamente!")
    print("=" * 70 + "\n")


if __name__ == "__main__":
    main()
