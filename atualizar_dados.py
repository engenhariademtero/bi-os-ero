#!/usr/bin/env python3
"""Lê o Report.xlsx sem dependências externas e gera dados anônimos para o BI."""

from __future__ import annotations

import argparse
import json
import math
import re
import shutil
import subprocess
import sys
import tempfile
import unicodedata
import zipfile
from datetime import datetime, timedelta
from pathlib import Path
from xml.etree import ElementTree as ET

ROOT = Path(__file__).resolve().parent
CONFIG = ROOT / "config.json"
OUTPUT = ROOT / "data.js"
LOG = ROOT / "sincronizacao.log"
NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main", "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}
REL_NS = {"p": "http://schemas.openxmlformats.org/package/2006/relationships"}

ALIASES = {
    "numero": ("NUMR OS", "NUMERO OS", "NUMERO DA OS", "N OS", "OS"),
    "instalacao": ("INSTALACAO", "LDAT", "CODSETOR", "COD SETOR"),
    "esquema": ("ESQUEMA", "ATIVIDADE", "DESCRICAO ESQUEMA"),
    "criacao": ("DATA CRIACAO", "CRIACAO", "DT CRIACAO"),
    "limite": ("DATA LIMITE", "LIMITE", "PRAZO"),
    "prevista": ("DATA PREVISTA", "PREVISAO", "DT PREVISTA"),
    "termino": ("DATA TERMINO", "TERMINO", "DATA FIM", "FIM", "TERMINO EXECUCAO"),
    "prioridade": ("PRIORID", "PRIORIDADE"),
    "situacao": ("SITUACAO", "STATUS", "SITUACAO DA OS"),
    "ativo": ("ATIVO", "EQUIPAMENTO"),
    "observacao": ("OBSERVACAO", "OBS", "DESCRICAO", "NOTA"),
}


def log(message: str) -> None:
    stamp = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    text = f"[{stamp}] {message}"
    print(text)
    with LOG.open("a", encoding="utf-8") as handle:
        handle.write(text + "\n")


def norm(value) -> str:
    if value is None:
        return ""
    text = unicodedata.normalize("NFKD", str(value))
    text = "".join(char for char in text if not unicodedata.combining(char))
    return re.sub(r"\s+", " ", re.sub(r"[^A-Za-z0-9]+", " ", text).strip()).upper()


def col_index(reference: str) -> int:
    letters = re.match(r"[A-Z]+", reference).group(0)
    result = 0
    for char in letters:
        result = result * 26 + ord(char) - 64
    return result - 1


def excel_date(value: float) -> datetime:
    return datetime(1899, 12, 30) + timedelta(days=value)


def is_date_format(code: str) -> bool:
    clean = re.sub(r'"[^\"]*"|\[[^\]]*\]|\\.', "", code.lower())
    return bool(re.search(r"(^|[^a-z])[dmyhs]+([^a-z]|$)", clean))


def read_xlsx(path: Path) -> list[list]:
    if not path.exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {path}")
    with zipfile.ZipFile(path) as book:
        shared = []
        if "xl/sharedStrings.xml" in book.namelist():
            root = ET.fromstring(book.read("xl/sharedStrings.xml"))
            for item in root.findall("m:si", NS):
                shared.append("".join(node.text or "" for node in item.findall(".//m:t", NS)))

        custom_formats = {}
        date_styles = set()
        if "xl/styles.xml" in book.namelist():
            root = ET.fromstring(book.read("xl/styles.xml"))
            for item in root.findall("m:numFmts/m:numFmt", NS):
                custom_formats[int(item.attrib["numFmtId"])] = item.attrib.get("formatCode", "")
            builtins = set(range(14, 23)) | set(range(27, 37)) | {45, 46, 47, 50, 57}
            for idx, xf in enumerate(root.findall("m:cellXfs/m:xf", NS)):
                num_id = int(xf.attrib.get("numFmtId", 0))
                if num_id in builtins or is_date_format(custom_formats.get(num_id, "")):
                    date_styles.add(idx)

        workbook = ET.fromstring(book.read("xl/workbook.xml"))
        first = workbook.find("m:sheets/m:sheet", NS)
        rel_id = first.attrib[f"{{{NS['r']}}}id"]
        rels = ET.fromstring(book.read("xl/_rels/workbook.xml.rels"))
        target = next(item.attrib["Target"] for item in rels.findall("p:Relationship", REL_NS) if item.attrib["Id"] == rel_id)
        sheet_path = "xl/" + target.lstrip("/") if not target.startswith("xl/") else target
        sheet = ET.fromstring(book.read(sheet_path))
        rows = []
        for row in sheet.findall("m:sheetData/m:row", NS):
            values = {}
            for cell in row.findall("m:c", NS):
                idx = col_index(cell.attrib["r"])
                kind = cell.attrib.get("t", "n")
                style = int(cell.attrib.get("s", 0))
                value_node = cell.find("m:v", NS)
                if kind == "inlineStr":
                    value = "".join(node.text or "" for node in cell.findall(".//m:t", NS))
                elif value_node is None:
                    value = None
                elif kind == "s":
                    value = shared[int(value_node.text)] if value_node.text else None
                elif kind in {"str", "e"}:
                    value = value_node.text
                elif kind == "b":
                    value = value_node.text == "1"
                else:
                    try:
                        number = float(value_node.text)
                        value = excel_date(number) if style in date_styles else (int(number) if number.is_integer() else number)
                    except (TypeError, ValueError):
                        value = value_node.text
                values[idx] = value
            width = max(values, default=-1) + 1
            rows.append([values.get(index) for index in range(width)])
        return rows


def parse_date(value):
    if value in (None, ""):
        return None
    if isinstance(value, datetime):
        return value
    if isinstance(value, (int, float)) and 20000 <= value <= 80000:
        return excel_date(float(value))
    text = str(value).strip()
    for fmt in ("%d/%m/%Y %H:%M:%S", "%d/%m/%Y", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            pass
    return None


def find_header(rows: list[list]) -> int:
    aliases = {norm(item) for item in ALIASES["numero"]}
    best = (-1, -1)
    for idx, row in enumerate(rows[:30]):
        keys = {norm(value) for value in row}
        score = len(keys & aliases) + int("SITUACAO" in keys) + int("INSTALACAO" in keys)
        if score > best[1]:
            best = (idx, score)
    if best[1] < 1:
        raise ValueError("Não foi possível localizar o cabeçalho do relatório.")
    return best[0]


def priority(value) -> str:
    key = norm(value)
    if "ESTRUT" in key: return "ESTRUTURANTE"
    if any(word in key for word in ("ALTA", "URGENT", "CRITIC")): return "ALTA"
    if "MEDIA" in key: return "MÉDIA"
    if "BAIXA" in key: return "BAIXA"
    return "SEM PRIORIDADE"


def present(value) -> bool:
    return value is not None and str(value).strip() != ""


def build_records(rows: list[list], reference: datetime) -> list[dict]:
    header_idx = find_header(rows)
    headers = [norm(value) for value in rows[header_idx]]
    mapping = {}
    for field, aliases in ALIASES.items():
        for alias in aliases:
            if norm(alias) in headers:
                mapping[field] = headers.index(norm(alias))
                break
    if "numero" not in mapping:
        raise ValueError("Coluna Número da OS não encontrada.")

    def get(row, field):
        idx = mapping.get(field)
        return row[idx] if idx is not None and idx < len(row) else None

    records = []
    for row in rows[header_idx + 1:]:
        number = get(row, "numero")
        if not present(number) or norm(number) in {norm(a) for a in ALIASES["numero"]}:
            continue
        installation = str(get(row, "instalacao") or "Não informado").strip().upper()
        creation = parse_date(get(row, "criacao"))
        deadline = parse_date(get(row, "limite")) or parse_date(get(row, "prevista"))
        termination = parse_date(get(row, "termino"))
        situation = norm(get(row, "situacao"))
        concluded = bool(re.search(r"EXECUT|CONCLUI|ENCERRAD|FINALIZ", situation))
        cancelled = "CANCEL" in situation
        running = not concluded and not cancelled and bool(re.search(r"EM EXECUCAO|ANDAMENTO|INICIAD", situation))
        planned = not concluded and not cancelled and not running and bool(re.search(r"PROGRAM|PLANEJ|PREVIST", situation))
        opened = not concluded and not cancelled
        status = "Concluída" if concluded else "Cancelada" if cancelled else "Atrasada" if deadline and deadline < reference else "Em execução" if running else "Programada" if planned else "Aberta no prazo"
        aging = max(0, (reference - creation).total_seconds() / 86400) if opened and creation else None
        delay = max(0, ((reference if opened else termination) - deadline).total_seconds() / 86400) if deadline and (opened or termination) else None
        pri = priority(get(row, "prioridade"))
        attention = "Encerrada"
        if opened:
            attention = "Média"
            if pri in {"ALTA", "ESTRUTURANTE"} or (delay or 0) > 30: attention = "Alta"
            if (pri == "ESTRUTURANTE" and (delay or 0) > 0) or (delay or 0) > 90: attention = "Crítica"
        compliance = "Não aplicável" if cancelled else "Sem data de prazo" if not deadline else "Sem data de término" if concluded and not termination else "No prazo" if concluded and termination <= deadline else "Fora do prazo" if concluded else "Vencida" if deadline < reference else "Em prazo"
        fields = [present(number), present(get(row, "instalacao")), present(get(row, "esquema")), creation is not None, deadline is not None, present(get(row, "situacao")), termination is not None, present(get(row, "ativo")), present(get(row, "observacao"))]
        complete = all([fields[1], fields[2], fields[3], fields[5], fields[8]]) and (not opened or fields[4]) and (not concluded or fields[6])
        records.append({"i": installation, "c": creation.strftime("%Y-%m") if creation else None, "t": termination.strftime("%Y-%m") if termination else None, "s": status, "p": pri, "a": round(aging, 1) if aging is not None else None, "n": attention, "u": compliance, "q": complete, "f": fields})
    return records


def run_git() -> None:
    if not (ROOT / ".git").exists():
        log("Repositório ainda não configurado. Dados locais foram atualizados.")
        return
    subprocess.run(["git", "pull", "--rebase"], cwd=ROOT, check=True)
    subprocess.run(["git", "add", "data.js"], cwd=ROOT, check=True)
    changed = subprocess.run(["git", "diff", "--cached", "--quiet"], cwd=ROOT).returncode != 0
    if not changed:
        log("Nenhuma alteração na base; envio não necessário.")
        return
    subprocess.run(["git", "commit", "-m", f"Atualiza base de OS em {datetime.now():%d/%m/%Y %H:%M}"], cwd=ROOT, check=True)
    subprocess.run(["git", "push"], cwd=ROOT, check=True)
    log("Site enviado ao GitHub Pages com sucesso.")


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--arquivo", help="Caminho alternativo do Report.xlsx")
    parser.add_argument("--sem-push", action="store_true", help="Gera os dados sem enviar ao GitHub")
    args = parser.parse_args()
    try:
        config = json.loads(CONFIG.read_text(encoding="utf-8"))
        source = Path(args.arquivo or config["arquivo_origem"])
        records = build_records(read_xlsx(source), datetime.now())
        payload = {"atualizado_em": datetime.now().isoformat(timespec="seconds"), "origem_atualizada_em": datetime.fromtimestamp(source.stat().st_mtime).isoformat(timespec="seconds"), "registros": records}
        temp = Path(tempfile.mkstemp(prefix="data_os_", suffix=".js", dir=ROOT)[1])
        temp.write_text("window.OS_DATA=" + json.dumps(payload, ensure_ascii=False, separators=(",", ":")) + ";\n", encoding="utf-8")
        temp.replace(OUTPUT)
        log(f"Base tratada: {len(records)} registros anônimos.")
        if not args.sem_push:
            run_git()
        return 0
    except Exception as error:
        log(f"ERRO: {type(error).__name__}: {error}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
