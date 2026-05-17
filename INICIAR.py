"""Menu principal amigável para usar o sistema PDF Amanda."""

from __future__ import annotations

import os
import subprocess
import sys
import traceback
from datetime import datetime
from pathlib import Path

import banco
from buscador import (
    busca_campo,
    busca_exata,
    busca_fuzzy,
    busca_regex,
    busca_semantica,
    busca_tabelas,
    carregar_indice,
    formatar_resultados,
)
from gerenciador_pastas import (
    FontePDF,
    adicionar_arquivo_config,
    adicionar_pasta_config,
    atualizar_pasta_principal,
    carregar_fontes_iniciais,
    fontes_para_processamento,
    pdfs_da_pasta,
    resumo_fontes,
    salvar_config_usuario,
    validar_arquivo_pdf,
    validar_pasta,
)
from pdf_system import INDICE, SAIDA, processar
from relatorios import exportar_busca

try:
    from colorama import Fore, Style, init as colorama_init

    colorama_init()
except Exception:  # pragma: no cover - fallback sem cores
    class _Cores:
        GREEN = YELLOW = RED = CYAN = RESET_ALL = BRIGHT = ""

    Fore = Style = _Cores()


BASE_DIR = Path(__file__).resolve().parent
ERROS_LOG = SAIDA / "erros.log"

for stream in (sys.stdout, sys.stderr):
    if hasattr(stream, "reconfigure"):
        stream.reconfigure(encoding="utf-8", errors="replace")


def registrar_erro(contexto: str, exc: BaseException) -> None:
    SAIDA.mkdir(exist_ok=True)
    with ERROS_LOG.open("a", encoding="utf-8") as f:
        f.write(f"\n[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {contexto}\n")
        f.write("".join(traceback.format_exception(type(exc), exc, exc.__traceback__)))
        f.write("\n")


def pausar() -> None:
    input("\nPressione ENTER para continuar...")


def limpar_tela() -> None:
    os.system("cls" if os.name == "nt" else "clear")


def gerar_ajuda() -> None:
    texto = """COMO USAR O SISTEMA PDF AMANDA
================================

PRIMEIRA VEZ:
  1. Clique duas vezes em INSTALAR.py
     (ou no terminal: python INSTALAR.py)
  Aguarde. O sistema instala tudo sozinho.

USO DO DIA A DIA:
  1. Clique duas vezes em INICIAR.bat (Windows)
     ou execute: python INICIAR.py
  2. Escolha o que deseja fazer digitando o número.

ADICIONAR NOVOS PDFs:
  Coloque o arquivo PDF nesta mesma pasta e execute o sistema normalmente.
  Também é possível escolher a opção 5 para usar outra pasta ou adicionar arquivos.

BUSCAR INFORMAÇÃO:
  Escolha opção 2 no menu e digite o que procura.
  Exemplos: artéria femoral, Dr. Carlos, R$ 2.325

SE ALGO DER ERRADO:
  O arquivo erros.log na pasta SAIDA tem os detalhes técnicos.
  Mostre esse arquivo para quem te ajudou a instalar.

ARQUIVOS GERADOS:
  Pasta SAIDA -> relatórios Word, Excel e texto de cada PDF.
  CONSOLIDADO.xlsx -> resumo de todos os documentos numa planilha.

CONSULTAR BASE DE DADOS (para listas de procedimentos):
  Quando o sistema detectar uma lista de registros repetidos,
  ele cria automaticamente uma base de dados consultável.
  Use a opção [6] do menu para buscar por nome, CPF, médico, etc.
  O arquivo DADOS_ESTRUTURADOS.xlsx na pasta SAIDA contém
  todos os registros em planilha organizada com várias abas.

PLANILHA DE PESQUISA (opção 7 do menu):
  Se você mantém uma planilha com prontuários de pacientes,
  o sistema pode completar automaticamente os dados de cada um.
  Você só precisa preencher ID, IMC e Tabagismo.

DASHBOARD VISUAL:
  Após processar os documentos, o sistema gera DASHBOARD.html
  na pasta SAIDA. Abra pela opção 3 do menu ou clique duas vezes
  no arquivo. Funciona em qualquer navegador.
"""
    (BASE_DIR / "AJUDA.txt").write_text(texto, encoding="utf-8")


class App:
    def __init__(self) -> None:
        self.fontes, self.config, self.mensagem_inicial = carregar_fontes_iniciais()
        gerar_ajuda()

    def cabecalho(self) -> None:
        limpar_tela()
        print(Fore.CYAN + "╔══════════════════════════════════════════════════╗")
        print("║         SISTEMA DE CONSULTA DE DOCUMENTOS        ║")
        print("║                   PDF AMANDA                     ║")
        print("╠══════════════════════════════════════════════════╣")
        print(f"║ Documentos carregados: {len(self.fontes):<3} PDFs                    ║")
        for nome, qtd in resumo_fontes(self.fontes)[:4]:
            nome_curto = nome[:31]
            print(f"║ • {nome_curto:<31} {qtd:>3} arquivo(s) ║")
        print("╚══════════════════════════════════════════════════╝" + Style.RESET_ALL)
        if self.mensagem_inicial:
            print(Fore.YELLOW + self.mensagem_inicial + Style.RESET_ALL)
            self.mensagem_inicial = ""

    def menu_principal(self) -> None:
        while True:
            try:
                self.cabecalho()
                print("\nO que você deseja fazer?\n")
                print("[1] Processar documentos (ler e extrair informações dos PDFs)")
                print("[2] Buscar informação nos documentos")
                print("[3] Ver relatórios gerados")
                print("[4] Ajuda")
                print("[5] Gerenciar pastas e arquivos")
                print("[6] Consultar base de dados estruturada")
                print("[7] Planilha de Pesquisa Clínica")
                print("[0] Sair")
                op = input("\nDigite o número da opção: ").strip()
                if op == "1":
                    self.processar_documentos()
                elif op == "2":
                    self.menu_busca()
                elif op == "3":
                    self.ver_relatorios()
                elif op == "4":
                    self.mostrar_ajuda()
                elif op == "5":
                    self.menu_gerenciar()
                elif op == "6":
                    self.menu_base_estruturada()
                elif op == "7":
                    self.menu_pesquisa_clinica()
                elif op == "0":
                    print("\nAté mais.")
                    break
                else:
                    print(Fore.YELLOW + "Opção inválida. Digite apenas o número da opção." + Style.RESET_ALL)
                    pausar()
            except Exception as exc:
                registrar_erro("Erro no menu principal", exc)
                print(Fore.RED + "Algo deu errado. O detalhe técnico foi salvo em SAIDA\\erros.log." + Style.RESET_ALL)
                pausar()

    def processar_documentos(self) -> None:
        if not self.fontes:
            print(Fore.YELLOW + "Nenhum PDF carregado. Use a opção 5 para escolher uma pasta ou arquivo." + Style.RESET_ALL)
            pausar()
            return
        print("\nVou ler os documentos e gerar os relatórios na pasta SAIDA.")
        resp = input("Deseja refazer relatórios já existentes? (s/N): ").strip().lower()
        reprocessar = resp == "s"
        try:
            processar(reprocessar=reprocessar, fontes=fontes_para_processamento(self.fontes))
            print(Fore.GREEN + "\nPronto. Os relatórios foram gerados/atualizados na pasta SAIDA." + Style.RESET_ALL)
        except Exception as exc:
            registrar_erro("Erro ao processar documentos", exc)
            print(Fore.RED + "Algo deu errado durante o processamento. Detalhes salvos em SAIDA\\erros.log." + Style.RESET_ALL)
        pausar()

    def _obter_indice(self):
        if not INDICE.exists():
            print(Fore.YELLOW + "Ainda não existe índice de busca. Vou processar os documentos primeiro." + Style.RESET_ALL)
            processar(fontes=fontes_para_processamento(self.fontes))
        return carregar_indice(INDICE)

    def menu_busca(self) -> None:
        try:
            indice = self._obter_indice()
        except Exception as exc:
            registrar_erro("Erro ao preparar busca", exc)
            print(Fore.RED + "Não consegui preparar a busca. Detalhes salvos em SAIDA\\erros.log." + Style.RESET_ALL)
            pausar()
            return

        ultimos = []
        ultimo_termo = ""
        while True:
            limpar_tela()
            print(Fore.CYAN + "BUSCAR INFORMAÇÃO NOS DOCUMENTOS" + Style.RESET_ALL)
            print("\n[1] Procurar por palavra ou frase")
            print("[2] Procurar mesmo com erro de digitação")
            print("[3] Procurar usando expressão avançada (regex)")
            print("[4] Procurar por relevância (busca inteligente)")
            print("[5] Procurar dentro das tabelas dos documentos")
            print("[6] Procurar em campo específico (ex: só no diagnóstico)")
            print("[7] Exportar resultado encontrado (TXT / Word / Excel)")
            print("[0] Voltar ao menu principal")
            op = input("\nDigite o número da opção: ").strip()
            if op == "0":
                return
            if op == "7":
                self.exportar_resultado(ultimos, ultimo_termo)
                continue
            termo = input("\nDigite o que você quer encontrar. Exemplo: artéria femoral\n> ").strip()
            if not termo:
                print("Nada foi digitado.")
                pausar()
                continue
            ultimo_termo = termo
            try:
                if op == "1":
                    ultimos = busca_exata(indice, termo)
                    if not ultimos:
                        print(Fore.YELLOW + "Não achei igualzinho. Vou tentar uma busca aproximada." + Style.RESET_ALL)
                        ultimos = busca_fuzzy(indice, termo, 0.70)
                elif op == "2":
                    ultimos = busca_fuzzy(indice, termo, 0.75)
                elif op == "3":
                    ultimos = busca_regex(indice, termo)
                elif op == "4":
                    ultimos = busca_semantica(indice, termo)
                elif op == "5":
                    ultimos = busca_tabelas(indice, termo)
                elif op == "6":
                    campo = input("Em qual campo? Exemplo: Diagnóstico, Paciente, CID\n> ").strip()
                    ultimos = busca_campo(indice, campo, termo)
                    ultimo_termo = f"{campo}_{termo}"
                else:
                    print("Opção inválida.")
                    pausar()
                    continue
                print(f"\nEncontrei {len(ultimos)} resultado(s).")
                print(formatar_resultados(ultimos))
                if ultimos and input("\nDeseja exportar esse resultado? (s/N): ").strip().lower() == "s":
                    self.exportar_resultado(ultimos, ultimo_termo)
            except Exception as exc:
                registrar_erro("Erro durante busca", exc)
                print(Fore.RED + "Não consegui concluir a busca. Detalhes salvos em SAIDA\\erros.log." + Style.RESET_ALL)
            pausar()

    def exportar_resultado(self, resultados, termo: str) -> None:
        if not resultados:
            print(Fore.YELLOW + "Ainda não há resultado para exportar." + Style.RESET_ALL)
            pausar()
            return
        formato = input("Escolha o formato: txt, docx ou xlsx\n> ").strip().lower() or "txt"
        if formato == "word":
            formato = "docx"
        if formato == "excel":
            formato = "xlsx"
        try:
            caminho = exportar_busca(resultados, termo or "busca", formato, SAIDA)
            print(Fore.GREEN + f"Resultado exportado para: {caminho}" + Style.RESET_ALL)
        except Exception as exc:
            registrar_erro("Erro ao exportar resultado", exc)
            print(Fore.RED + "Não consegui exportar. Detalhes salvos em SAIDA\\erros.log." + Style.RESET_ALL)

    def ver_relatorios(self) -> None:
        while True:
            limpar_tela()
            print(Fore.CYAN + "RELATÓRIOS GERADOS" + Style.RESET_ALL)
            print("\n[1] Listar arquivos da pasta SAIDA")
            print("[2] Mostrar caminho do Excel estruturado")
            print("[3] Mostrar caminho da planilha de pesquisa")
            print("[4] Abrir dashboard interativo (DASHBOARD.html)")
            print("[0] Voltar")
            op = input("\nDigite o número da opção: ").strip()
            if op == "0":
                return
            if op == "1":
                arquivos = sorted(SAIDA.glob("*"))
                if not arquivos:
                    print("Nenhum relatório foi gerado ainda. Use a opção 1 primeiro.")
                else:
                    print("\nRelatórios encontrados na pasta SAIDA:\n")
                    for arq in arquivos:
                        print(f"- {arq.name}")
                    print(f"\nPasta: {SAIDA}")
                pausar()
            elif op == "2":
                print(SAIDA / "DADOS_ESTRUTURADOS.xlsx")
                pausar()
            elif op == "3":
                print(SAIDA / "PESQUISA_CLINICA.xlsx")
                pausar()
            elif op == "4":
                self.abrir_dashboard()
                pausar()
            else:
                print("Opção inválida.")
                pausar()

    def abrir_dashboard(self) -> None:
        caminho = SAIDA / "DASHBOARD.html"
        if not caminho.exists():
            print("Dashboard ainda não foi gerado. Use a opção 1 para processar os documentos.")
            return
        try:
            if sys.platform.startswith("win"):
                os.startfile(caminho)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.run(["open", str(caminho)], check=False)
            else:
                subprocess.run(["xdg-open", str(caminho)], check=False)
            print(f"Dashboard aberto: {caminho}")
        except Exception as exc:
            registrar_erro("Erro ao abrir dashboard", exc)
            print(f"Não consegui abrir automaticamente. Abra este arquivo manualmente: {caminho}")

    def mostrar_ajuda(self) -> None:
        gerar_ajuda()
        print((BASE_DIR / "AJUDA.txt").read_text(encoding="utf-8"))
        pausar()

    def menu_base_estruturada(self) -> None:
        ultimos = []
        while True:
            limpar_tela()
            print(Fore.CYAN + "CONSULTAR BASE DE DADOS ESTRUTURADA" + Style.RESET_ALL)
            print("\n[1] Buscar por nome do paciente")
            print("[2] Buscar por CPF")
            print("[3] Buscar por prontuário")
            print("[4] Buscar por cirurgião")
            print("[5] Buscar por período (data início / data fim)")
            print("[6] Buscar por município")
            print("[7] Busca combinada (múltiplos filtros)")
            print("[8] Exportar resultado (TXT / Excel)")
            print("[0] Voltar")
            op = input("\nDigite o número da opção: ").strip()
            if op == "0":
                return
            if op == "8":
                self.exportar_estruturados(ultimos)
                pausar()
                continue
            filtros = {}
            if op == "1":
                filtros["nome"] = input("Digite o nome ou parte do nome. Exemplo: CRISTIANE\n> ").strip()
            elif op == "2":
                filtros["cpf"] = input("Digite o CPF, com ou sem pontos:\n> ").strip().replace(".", "").replace("-", "")
            elif op == "3":
                filtros["prontuario"] = input("Digite o prontuário:\n> ").strip()
            elif op == "4":
                filtros["cirurgiao"] = input("Digite o nome ou parte do nome do cirurgião:\n> ").strip()
            elif op == "5":
                filtros["data_inicio"] = input("Data inicial (AAAA-MM-DD):\n> ").strip()
                filtros["data_fim"] = input("Data final (AAAA-MM-DD):\n> ").strip()
            elif op == "6":
                filtros["municipio"] = input("Digite o município ou parte do nome:\n> ").strip()
            elif op == "7":
                filtros["nome"] = input("Nome do paciente (opcional):\n> ").strip()
                filtros["cpf"] = input("CPF (opcional):\n> ").strip().replace(".", "").replace("-", "")
                filtros["prontuario"] = input("Prontuário (opcional):\n> ").strip()
                filtros["cirurgiao"] = input("Cirurgião (opcional):\n> ").strip()
                filtros["municipio"] = input("Município (opcional):\n> ").strip()
                filtros = {k: v for k, v in filtros.items() if v}
            else:
                print("Opção inválida.")
                pausar()
                continue
            try:
                ultimos = banco.buscar_registros(**{k: v for k, v in filtros.items() if v})
                print(f"\nEncontrei {len(ultimos)} procedimento(s).")
                print(banco.imprimir_tabela_terminal(ultimos))
                if ultimos and input("\nDeseja exportar esse resultado? (s/N): ").strip().lower() == "s":
                    self.exportar_estruturados(ultimos)
            except Exception as exc:
                registrar_erro("Erro na consulta estruturada", exc)
                print("Não consegui consultar a base. Detalhes salvos em SAIDA\\erros.log.")
            pausar()

    def exportar_estruturados(self, registros) -> None:
        if not registros:
            print("Ainda não há resultados para exportar.")
            return
        formato = input("Formato: txt ou xlsx\n> ").strip().lower() or "txt"
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        if formato == "xlsx":
            from openpyxl import Workbook

            caminho = SAIDA / f"BUSCA_ESTRUTURADA_{ts}.xlsx"
            wb = Workbook()
            ws = wb.active
            ws.title = "Resultado"
            cols = ["prontuario", "nome", "cpf", "tipo_procedimento", "data_inicio", "cirurgiao", "municipio", "arquivo_origem"]
            ws.append(cols)
            for r in registros:
                ws.append([r.get(c, "") for c in cols])
            wb.save(caminho)
        else:
            caminho = SAIDA / f"BUSCA_ESTRUTURADA_{ts}.txt"
            caminho.write_text(banco.imprimir_tabela_terminal(registros, limite=100000), encoding="utf-8-sig")
        print(Fore.GREEN + f"Resultado exportado para: {caminho}" + Style.RESET_ALL)

    def menu_pesquisa_clinica(self) -> None:
        while True:
            limpar_tela()
            print(Fore.CYAN + "PLANILHA DE PESQUISA CLÍNICA" + Style.RESET_ALL)
            print("\n[1] Gerar planilha vazia com prontuários do banco")
            print("[2] Completar planilha existente")
            print("[3] Buscar prontuário específico")
            print("[4] Exportar seleção para planilha")
            print("[0] Voltar")
            op = input("\nDigite o número da opção: ").strip()
            if op == "0":
                return
            try:
                if op == "1":
                    caminho, encontrados, nao = banco.gerar_pesquisa_clinica()
                    print(f"Planilha criada: {caminho}")
                    print(f"{encontrados} linha(s) preenchidas. Não encontrados: {', '.join(nao) or 'nenhum'}")
                elif op == "2":
                    raw = input("Caminho da planilha (Enter para SAIDA\\PESQUISA_CLINICA.xlsx):\n> ").strip()
                    caminho = Path(raw.strip('"')) if raw else SAIDA / "PESQUISA_CLINICA.xlsx"
                    completados, nao = banco.completar_planilha_existente(caminho)
                    print(f"{completados} registros completados.")
                    if nao:
                        print(f"{len(nao)} prontuários não encontrados: {', '.join(nao[:30])}")
                elif op == "3":
                    pront = input("Digite o prontuário:\n> ").strip()
                    regs = banco.buscar_registros(prontuario=pront)
                    print(banco.imprimir_tabela_terminal(regs, limite=100))
                elif op == "4":
                    print("Use a opção [6] para buscar e exportar os resultados; o arquivo pode ser aberto no Excel.")
                else:
                    print("Opção inválida.")
            except Exception as exc:
                registrar_erro("Erro na planilha de pesquisa", exc)
                print("Não consegui concluir a operação. Detalhes salvos em SAIDA\\erros.log.")
            pausar()

    def recarregar_fontes(self) -> None:
        self.fontes, self.config, _ = carregar_fontes_iniciais()

    def menu_gerenciar(self) -> None:
        while True:
            limpar_tela()
            print(Fore.CYAN + "GERENCIAR PASTAS E ARQUIVOS" + Style.RESET_ALL)
            print("\n[1] Usar uma pasta diferente")
            print("[2] Adicionar outra pasta")
            print("[3] Adicionar arquivo(s) específico(s)")
            print("[4] Ver arquivos carregados atualmente")
            print("[5] Remover arquivo da lista atual")
            print("[6] Voltar ao menu principal")
            op = input("\nDigite o número da opção: ").strip()
            if op == "1":
                entrada = input("Digite ou cole o caminho da pasta:\n> ")
                pasta, msg = validar_pasta(entrada)
                print(msg)
                if pasta:
                    atualizar_pasta_principal(self.config, pasta)
                    self.recarregar_fontes()
            elif op == "2":
                entrada = input("Digite ou cole o caminho da pasta adicional:\n> ")
                pasta, msg = validar_pasta(entrada)
                print(msg)
                if pasta:
                    adicionar_pasta_config(self.config, pasta)
                    self.fontes.extend(pdfs_da_pasta(pasta, str(pasta), "pasta adicional"))
            elif op == "3":
                entrada = input("Cole um ou mais arquivos PDF. Pode separar por vírgula:\n> ")
                partes = [p.strip() for p in entrada.split(",") if p.strip()]
                for parte in partes:
                    arquivo, msg = validar_arquivo_pdf(parte)
                    print(msg)
                    if arquivo:
                        adicionar_arquivo_config(self.config, arquivo)
                        self.fontes.append(FontePDF(arquivo, str(arquivo.parent), "arquivo avulso"))
            elif op == "4":
                self.listar_fontes()
            elif op == "5":
                self.remover_fonte()
            elif op == "6":
                salvar_config_usuario(self.config)
                return
            else:
                print("Opção inválida.")
            pausar()

    def listar_fontes(self) -> None:
        if not self.fontes:
            print("Nenhum PDF carregado.")
            return
        print("\nArquivos carregados atualmente:\n")
        for i, fonte in enumerate(self.fontes, start=1):
            print(f"[{i}] {fonte.caminho.name}")
            print(f"    Origem: {fonte.origem}")

    def remover_fonte(self) -> None:
        self.listar_fontes()
        if not self.fontes:
            return
        escolha = input("\nDigite o número do arquivo para remover:\n> ").strip()
        if not escolha.isdigit():
            print("Digite apenas o número.")
            return
        idx = int(escolha) - 1
        if 0 <= idx < len(self.fontes):
            removido = self.fontes.pop(idx)
            print(f"Removido da lista atual: {removido.caminho.name}")
        else:
            print("Número inválido.")


def main() -> int:
    App().menu_principal()
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except KeyboardInterrupt:
        print("\nSistema fechado.")
        raise SystemExit(0)
