<#
.SYNOPSIS
    One-shot setup for the Mako Production System working folder on C:.
.DESCRIPTION
    Creates C:\MakoPS with this layout:
        01_Tools         inspection tooling (xl_inspect.py, Invoke-XlInspect.ps1, README.md)
        02_Source\2026   working copies of the PMC operation workbooks (originals on X: are not touched)
        03_Inspection    inspection_report.md plus one JSON per workbook
    Then runs the audit. Re-running is safe: files are overwritten, nothing is deleted.
.PARAMETER Root
    Target folder. Default C:\MakoPS.
.PARAMETER Source
    Folder holding the workbooks to copy. Default is the 2026 operation report folder on X:.
.PARAMETER SkipCopy
    Do not copy workbooks from X: (use when 02_Source\2026 is already populated).
.PARAMETER SkipRun
    Write files only, do not run the inspection.
.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Setup-MakoPS.ps1
#>
[CmdletBinding()]
param(
    [string]$Root = "C:\MakoPS",
    [string]$Source = "X:\13_Management\01_GM_Dashboard\01_Operation_Report\01_Excel\2026",
    [switch]$SkipCopy,
    [switch]$SkipRun
)

$ErrorActionPreference = "Stop"
$utf8NoBom = New-Object System.Text.UTF8Encoding $false
$utf8Bom   = New-Object System.Text.UTF8Encoding $true

function Write-TextFile([string]$Path, [string]$Content, [bool]$Bom) {
    $dir = Split-Path -Parent $Path
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }
    $enc = if ($Bom) { $utf8Bom } else { $utf8NoBom }
    [System.IO.File]::WriteAllText($Path, $Content, $enc)
    Write-Host ("  wrote {0}  ({1:N0} bytes)" -f $Path, (Get-Item $Path).Length)
}

Write-Host "Creating folder layout under $Root"
foreach ($sub in @("01_Tools", "02_Source\2026", "03_Inspection")) {
    New-Item -ItemType Directory -Force -Path (Join-Path $Root $sub) | Out-Null
}

Write-Host "Writing tooling"
$content_xl_inspect_py = @'
#!/usr/bin/env python3
"""
xl_inspect.py - Structural audit of Excel workbooks (.xlsx / .xlsm).

Purpose: inventory everything that makes a workbook fragile before rebuilding it
from scratch: external links, VBA, defined names, hidden sheets, cross-sheet
dependency graph, volatile functions, broken references, data connections.

Usage:
    python xl_inspect.py <folder-or-file> [more files...] [-o OUTPUT_DIR]

Outputs (in OUTPUT_DIR, default ./xl_inspect_out):
    inspection_report.md      human-readable report
    <workbook>.json           machine-readable detail per workbook

Dependencies: openpyxl (required), oletools (optional, for VBA code analysis).
"""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
import warnings
import zipfile
from urllib.parse import unquote
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path
from xml.etree import ElementTree as ET

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

try:
    import openpyxl
    from openpyxl.utils import get_column_letter
except ImportError:  # pragma: no cover
    sys.exit("openpyxl is required:  pip install openpyxl")

try:
    from oletools.olevba import VBA_Parser  # type: ignore

    HAVE_OLEVBA = True
except Exception:  # pragma: no cover
    HAVE_OLEVBA = False

try:
    import olefile  # type: ignore

    HAVE_OLEFILE = True
except Exception:  # pragma: no cover
    HAVE_OLEFILE = False


NS = {
    "m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}

VOLATILE = ("OFFSET(", "INDIRECT(", "NOW(", "TODAY(", "RAND(", "RANDBETWEEN(", "CELL(", "INFO(")
LOOKUPS = ("VLOOKUP(", "HLOOKUP(", "XLOOKUP(", "INDEX(", "MATCH(", "SUMIFS(", "SUMIF(", "COUNTIFS(", "SUMPRODUCT(", "GETPIVOTDATA(")
ERROR_VALUES = ("#REF!", "#N/A", "#VALUE!", "#DIV/0!", "#NAME?", "#NUM!", "#NULL!")
LARGE_FILE_MB = 60  # above this, cell scan runs in read-only mode

# 'Sheet Name'!A1  or  SheetName!A1  or  [1]Sheet!A1  or  'C:\path\[Book.xlsx]Sheet'!A1
SHEET_REF_RE = re.compile(r"(?<![\]#A-Za-z0-9_\.])(?:'((?:[^']|'')+)'|([A-Za-z0-9_\.]+))!")
EXT_INDEX_RE = re.compile(r"\[(\d+)\]")
VBA_PROC_RE = re.compile(r"^\s*(?:Public\s+|Private\s+|Friend\s+)?(?:Static\s+)?(Sub|Function|Property\s+(?:Get|Let|Set))\s+([A-Za-z_][A-Za-z0-9_]*)", re.I | re.M)
VBA_RISK_RE = re.compile(r"\b(Shell|Kill|CreateObject|GetObject|Application\.OnTime|SendKeys|Environ|FileCopy|RmDir|MkDir|Workbooks\.Open|ActiveWorkbook\.SaveAs|DisplayAlerts\s*=\s*False|On Error Resume Next|Sheets\([^)]*\)\.Delete|\.Delete)\b", re.I)


def human_size(n: int) -> str:
    for unit in ("B", "KB", "MB", "GB"):
        if n < 1024:
            return f"{n:.0f} {unit}" if unit == "B" else f"{n:.1f} {unit}"
        n /= 1024
    return f"{n:.1f} TB"


def parse_external_links(zf: zipfile.ZipFile) -> list[dict]:
    """Return list of external link targets in workbook order (index 1-based as used in formulas)."""
    links = []
    # workbook.xml.rels gives the order of externalLink parts as they appear in workbook.xml <externalReferences>
    try:
        wb_xml = ET.fromstring(zf.read("xl/workbook.xml"))
        wb_rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    except KeyError:
        return links
    rid_to_target = {r.get("Id"): r.get("Target") for r in wb_rels.findall("rel:Relationship", NS)}
    ext_refs = wb_xml.find("m:externalReferences", NS)
    if ext_refs is None:
        return links
    for idx, er in enumerate(ext_refs.findall("m:externalReference", NS), start=1):
        rid = er.get(f"{{{NS['r']}}}id")
        part = rid_to_target.get(rid, "")
        part_path = "xl/" + part if not part.startswith("/") else part.lstrip("/")
        rels_path = part_path.replace("externalLinks/", "externalLinks/_rels/") + ".rels"
        target = ""
        kind = "unknown"
        sheet_names: list[str] = []
        try:
            rels = ET.fromstring(zf.read(rels_path))
            for r in rels.findall("rel:Relationship", NS):
                target = r.get("Target", "")
                kind = r.get("Type", "").rsplit("/", 1)[-1]
        except KeyError:
            pass
        try:
            link_xml = ET.fromstring(zf.read(part_path))
            for sd in link_xml.iter(f"{{{NS['m']}}}sheetName"):
                sheet_names.append(sd.get("val", ""))
        except KeyError:
            pass
        links.append({"index": idx, "part": part_path, "target": target, "type": kind, "sheets_referenced": sheet_names})
    return links


def resolve_link_target(target: str, base_dir: Path) -> dict:
    """Best-effort check whether the external target exists on this machine."""
    t = unquote(target)
    if t.lower().startswith("file:///"):
        t = t[8:]
    elif t.lower().startswith("file://"):
        t = "\\\\" + t[7:]
    t = t.replace("/", os.sep)
    candidates = []
    if os.path.isabs(t) or re.match(r"^[A-Za-z]:", t) or t.startswith("\\\\"):
        candidates.append(Path(t))
    else:
        candidates.append(base_dir / t)
    for c in candidates:
        try:
            if c.exists():
                return {"resolved": True, "path": str(c)}
        except OSError:
            pass
    return {"resolved": False, "path": str(candidates[0]) if candidates else target}


def parse_connections(zf: zipfile.ZipFile) -> list[dict]:
    out = []
    if "xl/connections.xml" in zf.namelist():
        try:
            root = ET.fromstring(zf.read("xl/connections.xml"))
            for c in root.findall("m:connection", NS):
                entry = {"name": c.get("name"), "type": c.get("type"), "description": c.get("description", "")}
                db = c.find("m:dbPr", NS)
                if db is not None:
                    entry["connection"] = db.get("connection", "")[:300]
                    entry["command"] = db.get("command", "")[:300]
                out.append(entry)
        except ET.ParseError:
            pass
    return out


def parse_power_query(zf: zipfile.ZipFile) -> list[str]:
    names = []
    for n in zf.namelist():
        if n.startswith("customXml/item") and n.endswith(".xml"):
            try:
                data = zf.read(n)
                if b"DataMashup" in data:
                    names.append(n)
            except KeyError:
                pass
    return names


def count_parts(zf: zipfile.ZipFile, prefix: str) -> int:
    return sum(1 for n in zf.namelist() if n.startswith(prefix))


def sheet_drawing_counts(zf: zipfile.ZipFile) -> dict[str, dict]:
    """Map sheet part -> counts of charts / images / pivot tables via rels."""
    out: dict[str, dict] = {}
    names = set(zf.namelist())
    for n in names:
        m = re.match(r"xl/worksheets/_rels/(sheet\d+\.xml)\.rels$", n)
        if not m:
            continue
        sheet_part = m.group(1)
        counts = {"drawings": 0, "charts": 0, "images": 0, "pivot_tables": 0, "tables": 0, "comments": 0, "vml": 0}
        try:
            rels = ET.fromstring(zf.read(n))
        except ET.ParseError:
            continue
        for r in rels.findall("rel:Relationship", NS):
            typ = r.get("Type", "").rsplit("/", 1)[-1]
            target = r.get("Target", "")
            if typ == "drawing":
                counts["drawings"] += 1
                dpath = "xl/" + target.replace("../", "")
                drels = dpath.replace("drawings/", "drawings/_rels/") + ".rels"
                if drels in names:
                    try:
                        dr = ET.fromstring(zf.read(drels))
                        for rr in dr.findall("rel:Relationship", NS):
                            t2 = rr.get("Type", "").rsplit("/", 1)[-1]
                            if t2 == "chart":
                                counts["charts"] += 1
                            elif t2 == "image":
                                counts["images"] += 1
                    except ET.ParseError:
                        pass
            elif typ == "pivotTable":
                counts["pivot_tables"] += 1
            elif typ == "table":
                counts["tables"] += 1
            elif typ == "comments":
                counts["comments"] += 1
            elif typ == "vmlDrawing":
                counts["vml"] += 1
        out[sheet_part] = counts
    return out


def sheet_part_map(zf: zipfile.ZipFile) -> dict[str, str]:
    """Map sheet name -> sheetN.xml part name."""
    out = {}
    try:
        wb_xml = ET.fromstring(zf.read("xl/workbook.xml"))
        wb_rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    except KeyError:
        return out
    rid_to_target = {r.get("Id"): r.get("Target") for r in wb_rels.findall("rel:Relationship", NS)}
    sheets = wb_xml.find("m:sheets", NS)
    if sheets is None:
        return out
    for s in sheets.findall("m:sheet", NS):
        rid = s.get(f"{{{NS['r']}}}id")
        target = rid_to_target.get(rid, "")
        out[s.get("name")] = target.rsplit("/", 1)[-1]
    return out


def analyse_vba(path: Path) -> dict:
    result = {"present": False, "modules": [], "procedures": 0, "lines": 0, "auto_exec": [], "risk_hits": [], "error": None}
    if not HAVE_OLEVBA:
        result["error"] = "oletools not installed; VBA not analysed"
        return result
    try:
        vp = VBA_Parser(str(path))
    except Exception as exc:  # pragma: no cover
        result["error"] = f"olevba failed: {exc}"
        return result
    try:
        if not vp.detect_vba_macros():
            return result
        result["present"] = True
        for (_fn, _stream, vba_filename, code) in vp.extract_macros():
            code = code or ""
            lines = [ln for ln in code.splitlines() if ln.strip() and not ln.strip().startswith("Attribute ")]
            procs = [(m.group(1).split()[0], m.group(2)) for m in VBA_PROC_RE.finditer(code)]
            refs_sheets = sorted(set(re.findall(r"(?:Sheets|Worksheets)\(\"([^\"]+)\"\)", code)))
            refs_books = sorted(set(re.findall(r"Workbooks(?:\.Open)?\(\"([^\"]+)\"\)", code)))
            hardcoded_paths = sorted(set(re.findall(r"\"([A-Za-z]:\\[^\"]+|\\\\[^\"]+)\"", code)))
            risks = sorted(set(m.group(1) for m in VBA_RISK_RE.finditer(code)))
            module = {
                "module": vba_filename,
                "lines": len(lines),
                "procedures": [f"{k} {n}" for k, n in procs],
                "sheets_referenced": refs_sheets,
                "workbooks_referenced": refs_books,
                "hardcoded_paths": hardcoded_paths,
                "risk_keywords": risks,
            }
            result["modules"].append(module)
            result["procedures"] += len(procs)
            result["lines"] += len(lines)
            for _k, n in procs:
                if n.lower() in ("workbook_open", "auto_open", "workbook_beforeclose", "workbook_beforesave", "auto_close") or n.lower().startswith("worksheet_"):
                    result["auto_exec"].append(f"{vba_filename}:{n}")
            result["risk_hits"].extend(f"{vba_filename}:{r}" for r in risks)
        # Also record whether project is password protected
        try:
            for r in vp.analyze_macros():
                pass
        except Exception:
            pass
    finally:
        vp.close()
    return result


def scan_formula(f: str, own_sheet: str, ext_links: list[dict], stats: dict, deps: Counter, ext_use: Counter):
    fu = f.upper()
    stats["formulas"] += 1
    if any(v in fu for v in VOLATILE):
        stats["volatile"] += 1
    if any(l in fu for l in LOOKUPS):
        stats["lookups"] += 1
    if "#REF!" in fu:
        stats["ref_errors_in_formula"] += 1
    if fu.startswith("{=") or fu.startswith("{"):
        stats["array_formulas"] += 1
    ext_idx = [int(m.group(1)) for m in EXT_INDEX_RE.finditer(f)]
    if ext_idx:
        stats["external_link_formulas"] += 1
        for idx in ext_idx:
            ext_use[idx] += 1
    for m in SHEET_REF_RE.finditer(f):
        name = m.group(1) or m.group(2)
        if name is None:
            continue
        name = name.replace("''", "'")
        if "[" in name and "]" in name:
            # external workbook reference embedded in sheet ref - counted above
            continue
        if name != own_sheet:
            deps[(own_sheet, name)] += 1


def inspect_workbook(path: Path) -> dict:
    info: dict = {
        "file": str(path),
        "name": path.name,
        "size_bytes": path.stat().st_size,
        "size": human_size(path.stat().st_size),
        "modified": datetime.fromtimestamp(path.stat().st_mtime).strftime("%Y-%m-%d %H:%M"),
        "errors": [],
    }
    if not zipfile.is_zipfile(path):
        info["errors"].append("Not an OOXML zip (legacy .xls or .xlsb are not supported by this tool)")
        return info

    with zipfile.ZipFile(path) as zf:
        names = zf.namelist()
        info["has_vba_project"] = "xl/vbaProject.bin" in names
        info["external_links"] = parse_external_links(zf)
        for l in info["external_links"]:
            l.update(resolve_link_target(l["target"], path.parent))
        info["connections"] = parse_connections(zf)
        info["power_query_parts"] = parse_power_query(zf)
        info["part_counts"] = {
            "worksheets": count_parts(zf, "xl/worksheets/sheet"),
            "charts": count_parts(zf, "xl/charts/chart"),
            "drawings": count_parts(zf, "xl/drawings/drawing"),
            "pivot_caches": count_parts(zf, "xl/pivotCache/pivotCacheDefinition"),
            "pivot_tables": count_parts(zf, "xl/pivotTables/pivotTable"),
            "tables": count_parts(zf, "xl/tables/table"),
            "images": count_parts(zf, "xl/media/"),
            "query_tables": count_parts(zf, "xl/queryTables/"),
            "slicers": count_parts(zf, "xl/slicers/"),
            "external_link_parts": count_parts(zf, "xl/externalLinks/externalLink"),
        }
        info["uncompressed_bytes"] = sum(i.file_size for i in zf.infolist())
        largest = sorted(zf.infolist(), key=lambda i: i.file_size, reverse=True)[:8]
        info["largest_parts"] = [{"part": i.filename, "size": human_size(i.file_size)} for i in largest]
        per_sheet_draw = sheet_drawing_counts(zf)
        part_map = sheet_part_map(zf)

    # ---- VBA
    info["vba"] = analyse_vba(path) if info["has_vba_project"] else {"present": False, "modules": [], "procedures": 0, "lines": 0, "auto_exec": [], "risk_hits": [], "error": None}

    # ---- openpyxl pass
    big = info["size_bytes"] > LARGE_FILE_MB * 1024 * 1024
    try:
        wb = openpyxl.load_workbook(path, read_only=big, data_only=False, keep_vba=False, keep_links=True)
    except Exception as exc:
        info["errors"].append(f"openpyxl could not open workbook: {exc}")
        return info

    # defined names
    dnames = []
    try:
        dn_items = list(wb.defined_names.items()) if hasattr(wb.defined_names, "items") else [(d.name, d) for d in wb.defined_names.definedName]
    except Exception:
        dn_items = []
    for nm, d in dn_items:
        ref = d.attr_text or ""
        dnames.append({"name": nm, "scope": "workbook", "refers_to": ref, "broken": "#REF!" in ref, "external": bool(EXT_INDEX_RE.search(ref)) or "[" in ref, "hidden": bool(getattr(d, "hidden", False))})
    # sheet-scoped names
    if not big:
        for ws in wb.worksheets:
            try:
                items = list(ws.defined_names.items())
            except Exception:
                items = []
            for nm, d in items:
                ref = d.attr_text or ""
                dnames.append({"name": nm, "scope": ws.title, "refers_to": ref, "broken": "#REF!" in ref, "external": bool(EXT_INDEX_RE.search(ref)) or "[" in ref, "hidden": bool(getattr(d, "hidden", False))})
    info["defined_names"] = dnames

    sheets = []
    deps: Counter = Counter()
    ext_use: Counter = Counter()
    total = Counter()
    for ws in wb.worksheets:
        st = Counter()
        st_dict = {"formulas": 0, "volatile": 0, "lookups": 0, "ref_errors_in_formula": 0, "array_formulas": 0, "external_link_formulas": 0}
        values = 0
        formulas_sample: list[str] = []
        try:
            for row in ws.iter_rows():
                for c in row:
                    v = c.value
                    if v is None:
                        continue
                    values += 1
                    if isinstance(v, str) and v.startswith("="):
                        scan_formula(v, ws.title, info["external_links"], st_dict, deps, ext_use)
                        if len(formulas_sample) < 5 and ("[" in v or "!" in v):
                            formulas_sample.append(f"{c.coordinate}: {v[:120]}")
                    elif hasattr(v, "text"):  # ArrayFormula object
                        txt = str(getattr(v, "text", ""))
                        st_dict["array_formulas"] += 1
                        scan_formula("=" + txt if not txt.startswith("=") else txt, ws.title, info["external_links"], st_dict, deps, ext_use)
        except Exception as exc:
            info["errors"].append(f"cell scan failed on '{ws.title}': {exc}")
        entry = {
            "sheet": ws.title,
            "state": getattr(ws, "sheet_state", "visible"),
            "dims": ws.dimensions if not big else "n/a (read-only)",
            "max_row": ws.max_row,
            "max_col": ws.max_column,
            "cells_with_values": values,
            **st_dict,
        }
        if not big:
            entry["merged_ranges"] = len(ws.merged_cells.ranges)
            try:
                entry["data_validations"] = len(ws.data_validations.dataValidation)
            except Exception:
                entry["data_validations"] = 0
            try:
                entry["conditional_formats"] = sum(len(r.rules) for r in ws.conditional_formatting)
            except Exception:
                entry["conditional_formats"] = 0
            entry["tables"] = list(ws.tables.keys()) if hasattr(ws, "tables") else []
            entry["protected"] = bool(ws.protection.sheet)
            entry["freeze_panes"] = ws.freeze_panes
        part = part_map.get(ws.title)
        entry.update({k: v for k, v in per_sheet_draw.get(part, {}).items() if k in ("charts", "images", "pivot_tables", "comments")})
        entry["link_formula_samples"] = formulas_sample
        sheets.append(entry)
        for k, v in st_dict.items():
            total[k] += v
        total["cells_with_values"] += values
    # chartsheets
    for cs in getattr(wb, "chartsheets", []):
        sheets.append({"sheet": cs.title, "state": getattr(cs, "sheet_state", "visible"), "type": "chartsheet"})
    info["sheets"] = sheets
    info["totals"] = dict(total)
    info["cross_sheet_dependencies"] = [{"from": a, "to": b, "formulas": n} for (a, b), n in sorted(deps.items(), key=lambda kv: -kv[1])]
    known = {s["sheet"] for s in sheets}
    info["dangling_sheet_refs"] = sorted({b for (a, b) in deps if b not in known})
    for l in info["external_links"]:
        l["formulas_using"] = ext_use.get(l["index"], 0)
    wb.close()

    # ---- cached value error scan (data_only)
    err = Counter()
    err_by_sheet: dict[str, Counter] = defaultdict(Counter)
    try:
        wbv = openpyxl.load_workbook(path, read_only=True, data_only=True, keep_links=False)
        for ws in wbv.worksheets:
            for row in ws.iter_rows(values_only=True):
                for v in row:
                    if isinstance(v, str) and v in ERROR_VALUES:
                        err[v] += 1
                        err_by_sheet[ws.title][v] += 1
        wbv.close()
    except Exception as exc:
        info["errors"].append(f"cached value scan failed: {exc}")
    info["cached_errors"] = dict(err)
    info["cached_errors_by_sheet"] = {k: dict(v) for k, v in err_by_sheet.items()}
    return info


# ----------------------------------------------------------------------------- reporting

def md_table(headers: list[str], rows: list[list]) -> str:
    out = ["| " + " | ".join(headers) + " |", "|" + "|".join("---" for _ in headers) + "|"]
    for r in rows:
        out.append("| " + " | ".join(str(x) if x is not None else "" for x in r) + " |")
    return "\n".join(out)


def render_report(results: list[dict], base: Path) -> str:
    L: list[str] = []
    L.append(f"# Excel workbook inspection\n\nSource: `{base}`  \nGenerated: {datetime.now():%Y-%m-%d %H:%M}\n")

    # summary
    rows = []
    for r in results:
        if r.get("errors") and "sheets" not in r:
            rows.append([r["name"], r["size"], "ERROR", "", "", "", "", "", ""])
            continue
        hidden = sum(1 for s in r["sheets"] if s.get("state") != "visible")
        dead = sum(1 for l in r["external_links"] if not l.get("resolved"))
        rows.append([
            r["name"], r["size"], len(r["sheets"]), hidden,
            f"{len(r['external_links'])} ({dead} unresolved)",
            r["totals"].get("formulas", 0),
            r["totals"].get("external_link_formulas", 0),
            f"{r['vba']['procedures']} procs / {r['vba']['lines']} lines" if r["vba"]["present"] else ("bin present, unparsed" if r.get("has_vba_project") else "none"),
            sum(r["cached_errors"].values()),
        ])
    L.append("## Summary\n")
    L.append(md_table(["Workbook", "Size", "Sheets", "Hidden", "External links", "Formulas", "Ext-link formulas", "VBA", "Cached errors"], rows))
    L.append("")

    # cross-workbook link graph
    L.append("## Cross-workbook link map\n")
    names = {r["name"].lower(): r["name"] for r in results}
    graph_rows = []
    for r in results:
        for l in r.get("external_links", []):
            tgt = unquote(l["target"])
            tgt_name = tgt.replace("\\", "/").rsplit("/", 1)[-1]
            inset = "yes" if tgt_name.lower() in names else "NO (outside this set)"
            graph_rows.append([r["name"], f"[{l['index']}]", tgt[:110], "ok" if l.get("resolved") else "MISSING", inset, l.get("formulas_using", 0), ", ".join(l["sheets_referenced"][:6])])
    L.append(md_table(["From", "Idx", "Target", "On disk", "In this set", "Formulas", "Sheets referenced"], graph_rows) if graph_rows else "_No external links found._")
    L.append("")

    for r in results:
        L.append(f"\n---\n\n## {r['name']}\n")
        L.append(f"- Size {r['size']} on disk, {human_size(r.get('uncompressed_bytes', 0))} uncompressed. Modified {r['modified']}.")
        if r.get("errors"):
            for e in r["errors"]:
                L.append(f"- **Error:** {e}")
        if "sheets" not in r:
            continue
        pc = r["part_counts"]
        L.append(f"- Parts: {pc['worksheets']} worksheets, {pc['charts']} charts, {pc['pivot_tables']} pivot tables ({pc['pivot_caches']} caches), {pc['tables']} tables, {pc['images']} media files, {pc['query_tables']} query tables, {pc['slicers']} slicers.")
        L.append(f"- Data connections: {len(r['connections'])}; Power Query parts: {len(r['power_query_parts'])}.")
        L.append(f"- Largest parts: " + ", ".join(f"{p['part'].split('/')[-1]} {p['size']}" for p in r["largest_parts"][:5]))
        L.append("")

        # sheets
        L.append("### Sheets\n")
        srows = []
        for s in r["sheets"]:
            if s.get("type") == "chartsheet":
                srows.append([s["sheet"], s["state"], "chartsheet", "", "", "", "", "", "", "", ""])
                continue
            srows.append([
                s["sheet"], s["state"], s.get("dims", ""), s["cells_with_values"], s["formulas"],
                s["external_link_formulas"], s["volatile"], s["lookups"], s.get("data_validations", ""),
                s.get("conditional_formats", ""), f"{s.get('charts', 0)}c/{s.get('pivot_tables', 0)}p/{s.get('images', 0)}i",
            ])
        L.append(md_table(["Sheet", "State", "Used range", "Values", "Formulas", "Ext-link", "Volatile", "Lookups", "DV", "CF", "Charts/Pivots/Imgs"], srows))
        L.append("")

        # dependencies
        L.append("### Cross-sheet dependencies (formula count, top 40)\n")
        deps = r["cross_sheet_dependencies"][:40]
        L.append(md_table(["From sheet", "Reads from", "Formulas"], [[d["from"], d["to"], d["formulas"]] for d in deps]) if deps else "_None._")
        if r["dangling_sheet_refs"]:
            L.append(f"\n**Dangling sheet references (sheet no longer exists):** {', '.join(r['dangling_sheet_refs'])}")
        L.append("")

        # defined names
        dn = r["defined_names"]
        broken = [d for d in dn if d["broken"]]
        ext = [d for d in dn if d["external"] and not d["broken"]]
        L.append(f"### Defined names: {len(dn)} total, {len(broken)} broken (#REF!), {len(ext)} pointing to external workbooks\n")
        show = broken + ext
        if show:
            L.append(md_table(["Name", "Scope", "Refers to", "Status"], [[d["name"], d["scope"], d["refers_to"][:100], "BROKEN" if d["broken"] else "external"] for d in show[:60]]))
        L.append("")

        # VBA
        v = r["vba"]
        L.append("### VBA\n")
        if not v["present"]:
            if v.get("error"):
                L.append(f"_{v['error']}_")
            elif r.get("has_vba_project"):
                L.append("**vbaProject.bin is present but olevba found no parseable macros (corrupt, protected, or empty project). Open in the VBA editor to confirm.**")
            else:
                L.append("_No VBA project._")
        else:
            L.append(f"{len(v['modules'])} modules, {v['procedures']} procedures, {v['lines']} code lines.")
            if v["auto_exec"]:
                L.append(f"\n**Auto-executing procedures:** {', '.join(v['auto_exec'])}")
            if v["risk_hits"]:
                L.append(f"\n**Keywords to review:** {', '.join(sorted(set(v['risk_hits'])))}")
            L.append("")
            L.append(md_table(["Module", "Lines", "Procedures", "Sheets referenced", "Hard-coded paths"], [[m["module"], m["lines"], ", ".join(m["procedures"])[:200], ", ".join(m["sheets_referenced"])[:120], ", ".join(m["hardcoded_paths"])[:120]] for m in v["modules"]]))
        L.append("")

        # connections
        if r["connections"]:
            L.append("### Data connections\n")
            L.append(md_table(["Name", "Type", "Connection", "Command"], [[c.get("name"), c.get("type"), c.get("connection", "")[:100], c.get("command", "")[:100]] for c in r["connections"]]))
            L.append("")

        # cached errors
        if r["cached_errors"]:
            L.append("### Cached error values\n")
            L.append(md_table(["Sheet"] + list(ERROR_VALUES), [[sh] + [cnt.get(e, 0) for e in ERROR_VALUES] for sh, cnt in sorted(r["cached_errors_by_sheet"].items(), key=lambda kv: -sum(kv[1].values()))[:30]]))
            L.append("")

        # link formula samples
        samples = [(s["sheet"], f) for s in r["sheets"] for f in s.get("link_formula_samples", [])]
        if samples:
            L.append("### Sample cross-reference formulas\n")
            L.append("\n".join(f"- `{sh}` {f}" for sh, f in samples[:25]))
            L.append("")
    return "\n".join(L)


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("paths", nargs="+", help="folder(s) and/or workbook file(s)")
    ap.add_argument("-o", "--out", default="xl_inspect_out", help="output directory")
    args = ap.parse_args(argv)

    files: list[Path] = []
    for p in args.paths:
        pp = Path(p)
        if pp.is_dir():
            files.extend(sorted(x for x in pp.iterdir() if x.suffix.lower() in (".xlsx", ".xlsm", ".xltx", ".xltm", ".xls", ".xlsb") and not x.name.startswith("~$")))
        elif pp.exists():
            files.append(pp)
        else:
            print(f"warning: {p} not found", file=sys.stderr)
    if not files:
        sys.exit("no workbooks found")

    out = Path(args.out)
    out.mkdir(parents=True, exist_ok=True)
    results = []
    for f in files:
        print(f"inspecting {f.name} ...", flush=True)
        try:
            r = inspect_workbook(f)
        except Exception as exc:
            r = {"file": str(f), "name": f.name, "size": human_size(f.stat().st_size), "size_bytes": f.stat().st_size, "modified": "", "errors": [f"unhandled: {exc!r}"]}
        results.append(r)
        (out / (f.stem + ".json")).write_text(json.dumps(r, indent=2, default=str), encoding="utf-8")

    base = files[0].parent if len({f.parent for f in files}) == 1 else Path(os.path.commonpath([str(f) for f in files]))
    report = render_report(results, base)
    (out / "inspection_report.md").write_text(report, encoding="utf-8")
    print(f"\nreport written to {out / 'inspection_report.md'}")
    if not HAVE_OLEVBA:
        print("note: install oletools for VBA analysis:  pip install oletools")


if __name__ == "__main__":
    main()
'@
Write-TextFile (Join-Path $Root "01_Tools\xl_inspect.py") $content_xl_inspect_py $false
$content_Invoke_XlInspect_ps1 = @'
<#
.SYNOPSIS
    Runs xl_inspect.py against the PMC operation report workbooks.
.DESCRIPTION
    Installs openpyxl and oletools into the current Python if missing, then
    produces inspection_report.md and one JSON per workbook in the output folder.
.PARAMETER Source
    Folder holding the workbooks. Defaults to <root>\02_Source\2026 when it exists,
    otherwise the 2026 operation report folder on X:.
.PARAMETER Out
    Output folder for the report. Defaults to <root>\03_Inspection in the C:\MakoPS
    layout, otherwise <Source>\_inspection.
.EXAMPLE
    .\Invoke-XlInspect.ps1
    .\Invoke-XlInspect.ps1 -Source "X:\13_Management\01_GM_Dashboard\01_Operation_Report\01_Excel\2026"
#>
[CmdletBinding()]
param(
    [string]$Source = "",
    [string]$Out = ""
)

# Continue, not Stop: Windows PowerShell 5.1 turns native stderr output into terminating
# errors under Stop, and pip / openpyxl legitimately write warnings to stderr.
$ErrorActionPreference = "Continue"
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$script = Join-Path $here "xl_inspect.py"
$root = Split-Path -Parent $here

# Default layout: <root>\01_Tools (this script), <root>\02_Source\2026, <root>\03_Inspection
if ([string]::IsNullOrWhiteSpace($Source)) {
    $local = Join-Path $root "02_Source\2026"
    if (Test-Path $local) { $Source = $local }
    else { $Source = "X:\13_Management\01_GM_Dashboard\01_Operation_Report\01_Excel\2026" }
}
if (-not (Test-Path $Source)) { throw "Source folder not found: $Source" }
if ([string]::IsNullOrWhiteSpace($Out)) {
    if (Test-Path (Join-Path $root "02_Source")) { $Out = Join-Path $root "03_Inspection" }
    else { $Out = Join-Path $Source "_inspection" }
}

$python = Get-Command python -ErrorAction SilentlyContinue
if (-not $python) { $python = Get-Command py -ErrorAction SilentlyContinue }
if (-not $python) { throw "Python 3 not found on PATH. Install from python.org or the Microsoft Store." }
$py = $python.Source

function Test-PyModule([string]$Module) {
    # find_spec writes nothing to stderr, so this is safe on PowerShell 5.1
    & $py -c "import importlib.util, sys; sys.exit(0 if importlib.util.find_spec('$Module') else 1)"
    return ($LASTEXITCODE -eq 0)
}
foreach ($m in @("openpyxl", "oletools")) {
    if (-not (Test-PyModule $m)) {
        Write-Host "Installing $m ..."
        & $py -m pip install --quiet --disable-pip-version-check $m
        if (-not (Test-PyModule $m)) { throw "Could not install Python package '$m'. Install it manually:  $py -m pip install $m" }
    }
}

Write-Host "Inspecting workbooks in $Source"
& $py $script $Source -o $Out
if ($LASTEXITCODE -ne 0) { throw "xl_inspect.py failed with exit code $LASTEXITCODE" }

$report = Join-Path $Out "inspection_report.md"
Write-Host ""
Write-Host "Report: $report"
Write-Host "Attach the whole '$Out' folder (report + JSON files) to the session."
'@
Write-TextFile (Join-Path $Root "01_Tools\Invoke-XlInspect.ps1") $content_Invoke_XlInspect_ps1 $true
$content_README_md = @'
# tools

## xl_inspect.py

Structural audit of Excel workbooks (.xlsx / .xlsm) ahead of the Mako Production System rebuild.
For each workbook it reports:

- sheets with state (visible / hidden / veryHidden), used range, value and formula counts
- external links in formula index order, whether the target exists on disk, and how many formulas use each
- cross-sheet dependency graph and references to sheets that no longer exist
- defined names, flagging `#REF!` and external targets
- VBA modules, procedures, auto-executing events, hard-coded paths and keywords worth reviewing (needs `oletools`)
- data connections, Power Query parts, pivot caches, charts, tables, slicers
- volatile functions (OFFSET, INDIRECT, NOW, TODAY, RAND), lookups, array formulas
- cached error values (`#REF!`, `#N/A`, `#VALUE!`, `#DIV/0!`, `#NAME?`) per sheet

Legacy `.xls` and binary `.xlsb` are listed but not parsed; save them as `.xlsm` first.

### Run

Windows, from the repo root:

```powershell
.\tools\Invoke-XlInspect.ps1
.\tools\Invoke-XlInspect.ps1 -Source "X:\13_Management\01_GM_Dashboard\01_Operation_Report\01_Excel\2026"
```

Any platform:

```bash
pip install openpyxl oletools
python tools/xl_inspect.py <folder-or-files> -o <output-dir>
```

Output: `inspection_report.md` plus one `<workbook>.json` per file.
'@
Write-TextFile (Join-Path $Root "01_Tools\README.md") $content_README_md $false

# ---------------------------------------------------------------- copy workbooks
if (-not $SkipCopy) {
    $dest = Join-Path $Root "02_Source\2026"
    if (Test-Path $Source) {
        $files = Get-ChildItem -Path $Source -File | Where-Object {
            $_.Extension -match '^\.(xlsx|xlsm|xlsb|xls|xltm|xltx)$' -and $_.Name -notlike '~$*'
        }
        if ($files.Count -eq 0) {
            Write-Warning "No Excel files found in $Source"
        } else {
            Write-Host "Copying $($files.Count) workbook(s) from $Source to $dest"
            foreach ($f in $files) {
                Copy-Item -LiteralPath $f.FullName -Destination (Join-Path $dest $f.Name) -Force
                Write-Host ("  {0}  ({1:N1} MB)" -f $f.Name, ($f.Length / 1MB))
            }
        }
    } else {
        Write-Warning "Source folder not reachable: $Source. Copy the workbooks into $dest manually, then run 01_Tools\Invoke-XlInspect.ps1."
    }
}

# ---------------------------------------------------------------- run inspection
if (-not $SkipRun) {
    $runner = Join-Path $Root "01_Tools\Invoke-XlInspect.ps1"
    Write-Host ""
    Write-Host "Running inspection"
    & $runner -Source (Join-Path $Root "02_Source\2026") -Out (Join-Path $Root "03_Inspection")
    $report = Join-Path $Root "03_Inspection\inspection_report.md"
    if (Test-Path $report) {
        Write-Host ""
        Write-Host "Done. Report: $report"
        try { Start-Process explorer.exe (Join-Path $Root "03_Inspection") } catch { }
    }
}
