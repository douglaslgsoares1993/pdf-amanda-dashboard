"""Parser de relatórios institucionais com registros cirúrgicos repetitivos."""

from __future__ import annotations

import re
from dataclasses import dataclass, asdict
from datetime import datetime
from pathlib import Path
from typing import Any


SERVICOS_CONHECIDOS = [
    "CIRURGIA VASCULAR",
    "PEDIATRIA",
    "NEFROLOGIA",
    "CARDIOLOGIA",
    "CLINICA MEDICA",
    "ORTOPEDIA",
]


@dataclass
class RegistroEstruturado:
    prontuario: str = ""
    nome: str = ""
    cpf: str = ""
    dt_nascimento: str = ""
    idade: int | None = None
    sexo: str = ""
    municipio: str = ""
    servico: str = ""
    unidade_internacao: str = ""
    leito: str = ""
    tipo_atendimento: str = ""
    origem_paciente: str = ""
    cod_atendimento: str = ""
    data_atendimento: str = ""
    data_alta: str | None = None
    cirurgia: str = ""
    robotica: str = ""
    lateralidade: str = ""
    data_inicio: str = ""
    data_fim: str = ""
    duracao_minutos: int | None = None
    cirurgiao: str = ""
    especialidade: str = ""
    anestesista: str = ""
    centro_sala: str = ""
    arquivo_origem: str = ""
    tipo_procedimento: str = ""


def _limpar(texto: str) -> str:
    return re.sub(r"\s+", " ", (texto or "").strip())


def _parse_data_hora(valor: str) -> datetime | None:
    if not valor:
        return None
    for fmt in ("%d/%m/%Y %H:%M", "%d/%m/%Y"):
        try:
            return datetime.strptime(valor.strip(), fmt)
        except ValueError:
            pass
    return None


def _duracao(inicio: str, fim: str) -> int | None:
    di = _parse_data_hora(inicio)
    df = _parse_data_hora(fim)
    if not di or not df:
        return None
    return int((df - di).total_seconds() // 60)


def _tipo_procedimento(cirurgia: str, arquivo: str = "") -> str:
    texto = f"{cirurgia} {arquivo}".upper()
    if "ARTERIOGRAFIA" in texto:
        return "Arteriografia"
    if "AORTOGRAFIA" in texto:
        return "Aortografia"
    if "ANGIOPLASTIA" in texto:
        return "Angioplastia"
    return cirurgia.title()[:80] if cirurgia else Path(arquivo).stem


def _total_rodape(texto: str) -> int | None:
    m = re.search(r"Total\s+de\s+Procedimentos\s+Cirúrgicos:\s*(\d+)", texto, re.I)
    return int(m.group(1)) if m else None


def _linhas_validas(texto: str) -> list[str]:
    linhas = []
    for raw in texto.splitlines():
        linha = _limpar(raw)
        if not linha:
            continue
        if linha.startswith("═") or linha.startswith("--- PÁGINA"):
            continue
        if linha.startswith("UNIVERSIDADE DO ESTADO") or linha.startswith("HOSPITAL UNIVERSITÁRIO"):
            continue
        if linha.startswith("Relatório Detalhado") or linha.startswith("Período:"):
            continue
        if re.match(r"Página\s+\d+\s+de\s+\d+", linha):
            continue
        if linha.startswith("Prontuário Nome CPF"):
            continue
        if linha.startswith("Total de Procedimentos"):
            continue
        linhas.append(linha)
    return linhas


def _inicio_registro(linha: str) -> bool:
    return bool(re.match(r"^\d{5,}\s+", linha))


def _blocos(texto: str) -> list[list[str]]:
    linhas = _linhas_validas(texto)
    blocos: list[list[str]] = []
    atual: list[str] = []
    for linha in linhas:
        if _inicio_registro(linha):
            if atual:
                blocos.append(atual)
            atual = [linha]
        elif atual:
            atual.append(linha)
    if atual:
        blocos.append(atual)
    return blocos


def _separar_paciente(linhas_paciente: list[str]) -> dict[str, Any]:
    combinado = _limpar(" ".join(linhas_paciente))
    match = re.match(r"^(?P<prontuario>\d{5,})\s+(?P<nome>.+?)\s+(?:(?P<cpf>\d{11})\s+)?(?P<dn>\d{2}/\d{2}/\d{4})\s+(?P<idade>\d{1,3})\s+(?P<resto>.*)$", combinado)
    if not match:
        return {}
    dados = match.groupdict()
    nome = _limpar(dados["nome"])
    resto = _limpar(dados["resto"])
    sexo = ""
    sex_match = re.search(r"\b(?:0[67]|\d{2})\s+([MF])\b", resto)
    if sex_match:
        sexo = sex_match.group(1)
        resto = _limpar(resto[: sex_match.start()] + " " + resto[sex_match.end():])

    # Alguns nomes quebram depois do leito, antes do sexo: "... LEITO06 VERISSIMO 06 F".
    extra_match = re.search(r"(.+\d{3,4}-LEITO\d+)\s+([A-ZÁ-Ú ]{2,40})$", resto)
    if extra_match:
        extra = _limpar(extra_match.group(2))
        if extra not in {"RJ", "SP", "MG", "ES"}:
            nome = _limpar(f"{nome} {extra}")
            resto = _limpar(extra_match.group(1))

    leito = ""
    unidade = ""
    servico = ""
    municipio = resto
    leito_match = re.search(r"(\d{3,4}-LEITO\d+)$", resto)
    if leito_match:
        leito = leito_match.group(1)
        resto_sem_leito = resto[: leito_match.start()].strip()
    else:
        resto_sem_leito = resto
    unidade_match = re.search(r"(\d{4}\s+-\s+.+)$", resto_sem_leito)
    if unidade_match:
        unidade = unidade_match.group(1).strip()
        antes_unidade = resto_sem_leito[: unidade_match.start()].strip()
    else:
        antes_unidade = resto_sem_leito
    for serv in SERVICOS_CONHECIDOS:
        pos = antes_unidade.rfind(serv)
        if pos >= 0:
            servico = serv
            municipio = antes_unidade[:pos].strip()
            break
    return {
        "prontuario": dados["prontuario"],
        "nome": nome,
        "cpf": dados.get("cpf") or "",
        "dt_nascimento": dados["dn"],
        "idade": int(dados["idade"]),
        "sexo": sexo,
        "municipio": municipio,
        "servico": servico,
        "unidade_internacao": unidade,
        "leito": leito,
    }


def _parse_bloco(bloco: list[str], arquivo: str) -> RegistroEstruturado | None:
    try:
        idx_origem = next(i for i, l in enumerate(bloco) if l.startswith("Origem do Paciente:"))
        idx_cirurgia = next(i for i, l in enumerate(bloco) if l.startswith("Cirurgia:"))
        idx_cirurgiao = next(i for i, l in enumerate(bloco) if re.match(r"^Cirurgi[^:]{0,4}o:", l))
    except StopIteration:
        return None

    paciente = _separar_paciente(bloco[:idx_origem])
    if not paciente:
        return None

    origem_txt = " ".join(bloco[idx_origem:idx_cirurgia])
    cirurgia_txt = " ".join(bloco[idx_cirurgia:idx_cirurgiao])
    cirurgiao_txt = " ".join(bloco[idx_cirurgiao:])

    origem_match = re.search(
        r"Origem do Paciente:\s*(.*?)\s+Tipo Atendimento:\s*(.*?)\s+Cod\. Atendimento:\s*(\d+)\s+Data Atendimento:\s*(\d{2}/\d{2}/\d{4}\s+\d{2}:\d{2})\s+Data Alta:\s*(.*)$",
        origem_txt,
        re.I,
    )
    cirurgia_match = re.search(
        r"Cirurgia:\s*(.*?)\s+Rob.*?:\s*(.*?)\s+Lateralidade:\s*(.*?)\s+Data In.*?:\s*(\d{2}/\d{2}/\d{4}\s+\d{2}:\d{2})\s+Data Fim:\s*(\d{2}/\d{2}/\d{4}\s+\d{2}:\d{2})(.*)$",
        cirurgia_txt,
        re.I,
    )
    cirurgiao_match = re.search(r"Cirurgi[^:]{0,4}o:\s*(.*?)\s+Espec:\s*(.*?)\s+Anestesista:\s*(.*)$", cirurgiao_txt, re.I)
    if not origem_match or not cirurgia_match or not cirurgiao_match:
        return None

    cirurgia = _limpar(cirurgia_match.group(1))
    resto_cirurgia = _limpar(cirurgia_match.group(6))
    if resto_cirurgia and "STENT" in resto_cirurgia.upper() and "STENT" not in cirurgia.upper():
        cirurgia = _limpar(f"{cirurgia} {resto_cirurgia}")

    anestesia_resto = _limpar(cirurgiao_match.group(3))
    centro = ""
    anestesista = anestesia_resto
    centro_match = re.search(r"(CENTRO\s+.+?\s+SALA\s+\d+)$", anestesia_resto, re.I)
    if centro_match:
        centro = centro_match.group(1).strip()
        anestesista = anestesia_resto[: centro_match.start()].strip()

    registro = RegistroEstruturado(**paciente)
    registro.origem_paciente = _limpar(origem_match.group(1))
    registro.tipo_atendimento = _limpar(origem_match.group(2))
    registro.cod_atendimento = origem_match.group(3)
    registro.data_atendimento = origem_match.group(4)
    registro.data_alta = _limpar(origem_match.group(5)) or None
    registro.cirurgia = cirurgia
    registro.robotica = _limpar(cirurgia_match.group(2))
    registro.lateralidade = _limpar(cirurgia_match.group(3))
    registro.data_inicio = cirurgia_match.group(4)
    registro.data_fim = cirurgia_match.group(5)
    registro.duracao_minutos = _duracao(registro.data_inicio, registro.data_fim)
    registro.cirurgiao = _limpar(cirurgiao_match.group(1))
    registro.especialidade = _limpar(cirurgiao_match.group(2))
    registro.anestesista = anestesista
    registro.centro_sala = centro
    registro.arquivo_origem = arquivo
    registro.tipo_procedimento = _tipo_procedimento(cirurgia, arquivo)
    return registro


def parsear_registros(texto: str, arquivo_origem: str) -> dict[str, Any]:
    registros: list[dict[str, Any]] = []
    erros = 0
    for bloco in _blocos(texto):
        reg = _parse_bloco(bloco, arquivo_origem)
        if reg:
            registros.append(asdict(reg))
        else:
            erros += 1
    total = _total_rodape(texto)
    return {
        "arquivo": arquivo_origem,
        "registros": registros,
        "total_esperado": total,
        "total_extraido": len(registros),
        "erros": erros,
        "ok": total is None or total == len(registros),
    }
