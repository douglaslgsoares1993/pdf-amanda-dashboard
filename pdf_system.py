"""Sistema terminal de extração, busca e análise de PDFs médicos."""

from __future__ import annotations

import argparse
import hashlib
import sys
import traceback
from collections import Counter
from datetime import datetime
from pathlib import Path
from typing import Any

from buscador import (
    busca_campo,
    busca_exata,
    busca_fuzzy,
    busca_regex,
    busca_semantica,
    busca_tabelas,
    carregar_indice,
    construir_indice,
    formatar_resultados,
)
import banco
from config import CONFIG
from detector_tipo import TIPO_ESTRUTURADO, detectar_tipo_pdf
from extrator import extrair_pdf
from ia_analise import analisar_documento
from parser_registros import parsear_registros
from relatorios import (
    exportar_busca,
    garantir_saida,
    gerar_consolidado,
    gerar_relatorio_docx,
    gerar_relatorio_txt,
    gerar_texto_extraido,
    nome_base,
)

try:
    from tqdm import tqdm
except Exception:  # pragma: no cover - fallback para ambiente sem tqdm
    tqdm = None


BASE_DIR = Path(__file__).resolve().parent
SAIDA = garantir_saida(BASE_DIR)
INDICE = SAIDA / "INDICE_BUSCA.json"
ERROS_LOG = SAIDA / "erros.log"

for stream in (sys.stdout, sys.stderr):
    if hasattr(stream, "reconfigure"):
        stream.reconfigure(encoding="utf-8", errors="replace")


def log(msg: str) -> None:
    print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {msg}", flush=True)


def registrar_erro(contexto: str, exc: BaseException) -> None:
    """Grava detalhes técnicos sem despejar stack trace na tela."""
    ERROS_LOG.parent.mkdir(exist_ok=True)
    with ERROS_LOG.open("a", encoding="utf-8") as f:
        f.write(f"\n[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {contexto}\n")
        f.write("".join(traceback.format_exception(type(exc), exc, exc.__traceback__)))
        f.write("\n")


def listar_pdfs(arquivo: str | None = None) -> list[Path]:
    if arquivo:
        caminho = BASE_DIR / arquivo
        if not caminho.exists():
            caminho = Path(arquivo)
        return [caminho]
    return sorted(p for p in BASE_DIR.glob("*.pdf") if p.is_file())


def _fontes_padrao(arquivo: str | None = None) -> list[dict[str, str]]:
    return [
        {"caminho": str(pdf), "origem": str(pdf.parent), "tipo_origem": "pasta atual"}
        for pdf in listar_pdfs(arquivo)
    ]


def _preparar_fontes(fontes: list[dict[str, str]]) -> list[dict[str, str]]:
    vistos: set[str] = set()
    limpas: list[dict[str, str]] = []
    for fonte in fontes:
        caminho = Path(fonte["caminho"]).resolve()
        chave = str(caminho).lower()
        if chave in vistos:
            continue
        vistos.add(chave)
        limpas.append({
            "caminho": str(caminho),
            "origem": fonte.get("origem") or str(caminho.parent),
            "tipo_origem": fonte.get("tipo_origem", "pasta"),
        })
    nomes = Counter(Path(f["caminho"]).name.lower() for f in limpas)
    for fonte in limpas:
        caminho = Path(fonte["caminho"])
        base = nome_base(caminho.name)
        if nomes[caminho.name.lower()] > 1:
            curto = hashlib.sha1(str(caminho).encode("utf-8", errors="ignore")).hexdigest()[:4]
            base = f"{base}_{curto}"
        fonte["saida_base"] = base
    return limpas


def saidas_pdf_existem(fonte: dict[str, str]) -> bool:
    base = fonte.get("saida_base") or nome_base(Path(fonte["caminho"]).name)
    return (
        (SAIDA / f"{base}_TEXTO.txt").exists()
        and (SAIDA / f"{base}_RELATORIO.txt").exists()
        and (SAIDA / f"{base}_RELATORIO.docx").exists()
    )


def _iterar_fontes(fontes: list[dict[str, str]]):
    if tqdm:
        return tqdm(fontes, desc="Processando PDFs", unit="pdf")
    return fontes


def processar(
    arquivo: str | None = None,
    reprocessar: bool = False,
    fontes: list[dict[str, str]] | None = None,
) -> list[dict[str, Any]]:
    """Processa PDFs e atualiza relatórios, consolidado e índice."""
    log("Iniciando processamento dos PDFs.")
    log(f"Provedor de IA selecionado: {CONFIG.provider or 'indisponível'}")
    detectadas = ", ".join(k for k, v in CONFIG.available_keys.items() if v) or "nenhuma"
    log(f"Chaves detectadas: {detectadas}")

    fontes_processar = _preparar_fontes(fontes if fontes is not None else _fontes_padrao(arquivo))
    if not fontes_processar:
        log("Nenhum PDF encontrado.")
        return []

    resultados: list[dict[str, Any]] = []
    analises: dict[str, dict[str, Any]] = {}
    estruturados_por_tipo: dict[str, int] = {}
    houve_estruturado = False
    if reprocessar:
        banco.inicializar_banco(True)
    else:
        banco.inicializar_banco(False)

    for fonte in _iterar_fontes(fontes_processar):
        pdf = Path(fonte["caminho"])
        if tqdm:
            try:
                _iterar_fontes  # Mantém linters quietos; tqdm atualiza o item sozinho.
            except Exception:
                pass
        if not pdf.exists():
            log(f"Arquivo não encontrado: {pdf}")
            continue
        try:
            tipo_info = detectar_tipo_pdf(pdf)
            tipo_doc = tipo_info.get("tipo")
            log(f"Tipo detectado em {pdf.name}: {tipo_doc}")
            if saidas_pdf_existem(fonte) and not reprocessar:
                log(f"Saídas já existem para {pdf.name}; lendo novamente para consolidado e índice.")
            else:
                log(f"Extraindo: {pdf.name}")

            doc = extrair_pdf(pdf)
            doc["id"] = str(pdf.resolve())
            doc["origem"] = fonte.get("origem", str(pdf.parent))
            doc["tipo_origem"] = fonte.get("tipo_origem", "pasta")
            doc["saida_base"] = fonte.get("saida_base")

            if doc.get("erro"):
                log(f"Aviso em {pdf.name}: {doc['erro']}")
            log(f"Motor usado em {pdf.name}: {doc.get('motor')} (score {doc.get('score')})")
            for campo, valores in doc.get("campos", {}).items():
                if not valores:
                    log(f"{pdf.name}: campo não encontrado - {campo}")

            analise = analisar_documento(doc.get("texto_completo", ""), doc.get("campos", {}))
            if analise.get("observacao_ia"):
                log(f"{pdf.name}: {analise['observacao_ia']}")
            analises[doc["arquivo"]] = analise
            resultados.append(doc)

            gerar_texto_extraido(doc, SAIDA, reprocessar)
            gerar_relatorio_txt(doc, analise, SAIDA, reprocessar)
            gerar_relatorio_docx(doc, analise, SAIDA, reprocessar)

            if tipo_doc == TIPO_ESTRUTURADO:
                estruturado = parsear_registros(doc.get("texto_completo", ""), str(pdf.resolve()))
                inseridos = banco.inserir_registros(estruturado["registros"])
                houve_estruturado = True
                tipo_proc = estruturado["registros"][0]["tipo_procedimento"] if estruturado["registros"] else pdf.stem
                estruturados_por_tipo[tipo_proc] = estruturados_por_tipo.get(tipo_proc, 0) + estruturado["total_extraido"]
                esperado = estruturado.get("total_esperado")
                status = "✔" if estruturado.get("ok") else "⚠"
                if esperado is not None:
                    log(f"{status} {tipo_proc}: {estruturado['total_extraido']} extraídos / {esperado} esperados ({inseridos} novos no banco)")
                else:
                    log(f"{status} {tipo_proc}: {estruturado['total_extraido']} registros extraídos ({inseridos} novos no banco)")
        except Exception as exc:
            registrar_erro(f"Falha ao processar {pdf}", exc)
            log(f"Algo deu errado ao processar o arquivo {pdf.name}. O sistema continuou com os outros arquivos.")

    if resultados:
        gerar_consolidado(resultados, analises, SAIDA)
        construir_indice(resultados, INDICE)
        log(f"Consolidado atualizado: {SAIDA / 'CONSOLIDADO.xlsx'}")
        log(f"Índice de busca atualizado: {INDICE}")
    if houve_estruturado:
        xlsx = banco.gerar_excel_estruturado()
        dashboard = banco.gerar_dashboard_html()
        prontuarios_teste = [
            2577507, 2504626, 2577355, 2574108, 1934547, 2571753, 2541803, 1902136,
            2544295, 2565233, 2523984, 220664, 2549620, 2560230, 2032289, 2559481,
            2546357, 1533129, 2562699, 2546212, 2455133, 2405095, 2558345, 2555576,
            2553894, 2545958,
        ]
        pesquisa, encontrados, nao_encontrados = banco.gerar_pesquisa_clinica(prontuarios_teste)
        for tipo, total in sorted(estruturados_por_tipo.items()):
            log(f"✔ {tipo}: {total} registros")
        log(f"✔ Total no banco: {banco.total_banco()} registros")
        log(f"✔ DADOS_ESTRUTURADOS.xlsx gerado com 8 abas: {xlsx}")
        log(f"✔ Dashboard atualizado: {dashboard}")
        log(f"✔ PESQUISA_CLINICA.xlsx atualizada: {pesquisa} ({encontrados} registros completados; não encontrados: {', '.join(nao_encontrados) or 'nenhum'})")
    log(f"Processamento concluído. PDFs processados: {len(resultados)}")
    return resultados


def garantir_indice(fontes: list[dict[str, str]] | None = None) -> dict[str, Any]:
    if not INDICE.exists():
        log("Índice não encontrado. Processando PDFs antes da busca.")
        processar(fontes=fontes)
    return carregar_indice(INDICE)


def executar_busca(tipo: str, termo: str, indice: dict[str, Any], extra: str = "") -> list[dict[str, Any]]:
    if tipo == "exata":
        return busca_exata(indice, termo)
    if tipo == "fuzzy":
        try:
            minimo = float(extra) if extra else 0.75
        except ValueError:
            minimo = 0.75
        return busca_fuzzy(indice, termo, minimo)
    if tipo == "regex":
        return busca_regex(indice, termo)
    if tipo == "semantica":
        return busca_semantica(indice, termo)
    if tipo == "tabelas":
        return busca_tabelas(indice, termo)
    if tipo == "campo":
        return busca_campo(indice, extra or termo, "" if extra else termo)
    return []


def perguntar_exportacao(resultados: list[dict[str, Any]], termo: str) -> None:
    if not resultados:
        return
    try:
        resp = input("Exportar resultado? (s/N): ").strip().lower()
    except EOFError:
        log("Entrada interativa indisponível; exportação da busca ignorada.")
        return
    if resp != "s":
        return
    try:
        formato = input("Formato (txt/docx/xlsx): ").strip().lower() or "txt"
    except EOFError:
        formato = "txt"
    try:
        caminho = exportar_busca(resultados, termo, formato, SAIDA)
        log(f"Busca exportada para: {caminho}")
    except Exception as exc:
        registrar_erro("Falha ao exportar busca", exc)
        log("Não consegui exportar a busca. Os detalhes foram salvos em SAIDA\\erros.log.")


def busca_direta(termo: str, fontes: list[dict[str, str]] | None = None) -> list[dict[str, Any]]:
    indice = garantir_indice(fontes)
    resultados = busca_exata(indice, termo)
    if not resultados:
        log("Nenhum literal encontrado; tentando busca aproximada.")
        resultados = busca_fuzzy(indice, termo, 0.70)
    print(formatar_resultados(resultados))
    perguntar_exportacao(resultados, termo)
    return resultados


def menu_busca() -> None:
    indice = garantir_indice()
    ultimos: list[dict[str, Any]] = []
    ultimo_termo = ""
    while True:
        print(
            "\n[1] Busca exata\n[2] Busca aproximada (fuzzy)\n[3] Busca por regex\n"
            "[4] Busca semântica\n[5] Busca em tabelas\n[6] Buscar em campo específico\n"
            "[7] Exportar último resultado para TXT / DOCX / XLSX\n[0] Sair"
        )
        op = input("Opção: ").strip()
        if op == "0":
            break
        if op == "7":
            perguntar_exportacao(ultimos, ultimo_termo)
            continue
        termo = input("Termo/consulta: ").strip()
        ultimo_termo = termo
        if op == "1":
            ultimos = busca_exata(indice, termo)
        elif op == "2":
            minimo = input("Similaridade mínima (padrão 0.75): ").strip()
            ultimos = executar_busca("fuzzy", termo, indice, minimo)
        elif op == "3":
            ultimos = busca_regex(indice, termo)
        elif op == "4":
            ultimos = busca_semantica(indice, termo)
        elif op == "5":
            ultimos = busca_tabelas(indice, termo)
        elif op == "6":
            campo = input("Campo (ex: Diagnóstico, Paciente, CID): ").strip()
            ultimos = busca_campo(indice, campo, termo)
            ultimo_termo = f"{campo}_{termo}"
        else:
            print("Opção inválida.")
            continue
        print(formatar_resultados(ultimos))
        perguntar_exportacao(ultimos, ultimo_termo)


def main() -> int:
    parser = argparse.ArgumentParser(description="Extração, busca e análise de PDFs médicos.")
    parser.add_argument("--arquivo", help="Processa apenas um PDF específico.")
    parser.add_argument("--buscar", nargs="?", const=True, help="Abre menu de busca ou busca termo direto.")
    parser.add_argument("--reprocessar", action="store_true", help="Força reprocessamento das saídas.")
    args = parser.parse_args()

    if args.buscar is True:
        menu_busca()
    elif isinstance(args.buscar, str):
        busca_direta(args.buscar)
    else:
        processar(args.arquivo, args.reprocessar)
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except KeyboardInterrupt:
        print("\nOperação interrompida pelo usuário.")
        raise SystemExit(130)
