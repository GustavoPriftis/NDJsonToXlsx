#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import argparse, json, os, sys
import json, datetime
from decimal import Decimal

try:
    import xlsxwriter
except Exception:
    xlsxwriter = None

def flatten(obj, prefix="", out=None):
    if out is None:
        out = {}
    if obj is None:
        if prefix:
            out[prefix] = None
        return out
    if isinstance(obj, (str, int, float, bool)):
        if prefix:
            out[prefix] = obj
        return out
    if isinstance(obj, list):
        for i, v in enumerate(obj):
            key = f"{prefix}[{i}]" if prefix else f"[{i}]"
            flatten(v, key, out)
        return out
    if isinstance(obj, dict):
        for k, v in obj.items():
            key = f"{prefix}.{k}" if prefix else k
            flatten(v, key, out)
        return out
    if prefix:
        out[prefix] = str(obj)
    return out

def collect_headers(path):
    headers = set()
    total = 0
    invalid = 0
    with open(path, "r", encoding="utf-8") as f:
        for line in f:
            s = line.strip()
            if not s:
                continue
            total += 1
            try:
                flat = flatten(json.loads(s))
                headers.update(flat.keys())
            except json.JSONDecodeError:
                invalid += 1
    return sorted(headers), total, invalid

def write_xlsx(path_in, path_out, headers):
    wb = xlsxwriter.Workbook(path_out)
    ws = wb.add_worksheet("Dados")
    head_fmt = wb.add_format({"bold": True})
    for col, h in enumerate(headers):
        ws.write(0, col, h, head_fmt)

    col_widths = [max(8, len(h) + 2) for h in headers]

    row_idx = 1
    written = 0
    skipped_invalid = 0

    with open(path_in, "r", encoding="utf-8") as f:
        for line in f:
            s = line.strip()
            if not s:
                continue
            try:
                flat = flatten(json.loads(s))
            except json.JSONDecodeError:
                skipped_invalid += 1
                continue

            for col, h in enumerate(headers):
                raw = flat.get(h, "")
                val = to_excel_value(raw)
                ws.write(row_idx, col, val)
                l = len(str(val)) if val is not None else 0
                if l + 2 > col_widths[col]:
                    col_widths[col] = min(l + 2, 62)  
            row_idx += 1
            written += 1

    for c, w in enumerate(col_widths):
        ws.set_column(c, c, w)

    wb.close()
    return written, skipped_invalid

def to_excel_value(v):
    if v is None or isinstance(v, (str, int, float, bool)):
        return v
    if isinstance(v, (datetime.date, datetime.datetime, datetime.time)):
        return v.isoformat()
    if isinstance(v, Decimal):
        try:
            return float(v)
        except Exception:
            return str(v)
    try:
        return json.dumps(v, ensure_ascii=False)
    except Exception:
        return str(v)

def main():
    ap = argparse.ArgumentParser(description="Converte NDJSON (um JSON por linha) em Excel (.xlsx).")
    ap.add_argument("input", help="Arquivo NDJSON de entrada.")
    ap.add_argument("output", help="Arquivo .xlsx de saída.")
    args = ap.parse_args()

    if not os.path.isfile(args.input):
        print("Arquivo de entrada não existe.", file=sys.stderr)
        sys.exit(1)
    if not args.output.lower().endswith(".xlsx"):
        print("A saída deve terminar com .xlsx", file=sys.stderr)
        sys.exit(1)

    headers, total, invalid_first = collect_headers(args.input)
    if not headers:
        print("Nenhum campo encontrado em linhas válidas.", file=sys.stderr)
        sys.exit(1)

    written, invalid_second = write_xlsx(args.input, args.output, headers)

    print(f"Linhas totais lidas: {total}")
    print(f"Inválidas ignoradas: {invalid_first} (descoberta de colunas) + {invalid_second} (gravação)")
    print(f"Gerado: {args.output} com {written} linhas e {len(headers)} colunas.")

if __name__ == "__main__":
    main()