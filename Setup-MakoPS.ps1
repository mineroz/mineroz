<#
.SYNOPSIS
    One-shot setup for the Mako Production System working folder on C:.
.DESCRIPTION
    Creates C:\MakoPS with this layout:
        01_Tools         tooling: xl_inspect.py, PythonEnv.ps1, Invoke-XlInspect.ps1, Build-MPS.ps1,
                         Export-DailyReport.ps1, README.md
        01_Tools\mps     MPS generator: build_workbook.py, build_ddl.py, validate_workbook.py,
                         schema\mps_schema.json
        02_Source\2026   working copies of the PMC operation workbooks (originals on X: are not touched)
        03_Inspection    inspection_report.md plus one JSON per workbook
        04_MPS           MPS_<Year>.xlsx (blank) and MPS_<Year>_DEMO.xlsx (synthetic data)
    Then runs the audit and builds the MPS workbooks. Re-running is safe: files are overwritten,
    nothing is deleted. Python 3 is resolved once (installed for the current user when it is
    missing); when that fails the audit and the build are skipped with a warning and the setup
    still finishes, so the tooling is in place for a later run.
    Run from a normal PowerShell window: mapped drives such as X: are not visible in an
    elevated (Run as administrator) session.
.PARAMETER Root
    Target folder. Default C:\MakoPS.
.PARAMETER Source
    Folder holding the workbooks to copy. Default is the 2026 operation report folder on X:.
.PARAMETER Year
    Reporting year of the MPS workbooks. Default 2026.
.PARAMETER SkipCopy
    Do not copy workbooks from X: (use when 02_Source\2026 is already populated).
.PARAMETER SkipRun
    Do not run the inspection.
.PARAMETER SkipBuild
    Do not build the MPS workbooks.
.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Setup-MakoPS.ps1
    powershell -ExecutionPolicy Bypass -File .\Setup-MakoPS.ps1 -SkipCopy -SkipRun
#>
[CmdletBinding()]
param(
    [string]$Root = "C:\MakoPS",
    [string]$Source = "X:\13_Management\01_GM_Dashboard\01_Operation_Report\01_Excel\2026",
    [int]$Year = 2026,
    [switch]$SkipCopy,
    [switch]$SkipRun,
    [switch]$SkipBuild
)

$ErrorActionPreference = "Stop"
# Absolute root: the .NET file writer below resolves relative paths against the process directory, not the PowerShell location
$Root = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Root)
$utf8NoBom = New-Object System.Text.UTF8Encoding $false
$utf8Bom   = New-Object System.Text.UTF8Encoding $true

function Write-TextFile([string]$Path, [string]$Content, [bool]$Bom) {
    $dir = Split-Path -Parent $Path
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }
    $enc = if ($Bom) { $utf8Bom } else { $utf8NoBom }
    [System.IO.File]::WriteAllText($Path, $Content, $enc)
    Write-Host ("  wrote {0}  ({1:N0} bytes)" -f $Path, (Get-Item $Path).Length)
}

$elevated = $false
try {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $elevated = (New-Object Security.Principal.WindowsPrincipal $identity).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
} catch { }
if ($elevated) {
    Write-Warning "This PowerShell session is elevated (Run as administrator). Mapped drives such as X: are not visible here; if the copy step fails, re-run from a normal PowerShell window."
}

Write-Host "Creating folder layout under $Root"
foreach ($sub in @("01_Tools", "01_Tools\mps\schema", "02_Source\2026", "03_Inspection", "04_MPS")) {
    New-Item -ItemType Directory -Force -Path (Join-Path $Root $sub) | Out-Null
}

Write-Host "Writing tooling"
$content_tools_xl_inspect_py = @'
#!/usr/bin/env python3
"""
xl_inspect.py - Structural audit of Excel workbooks (.xlsx / .xlsm).

Purpose: inventory everything that makes a workbook fragile before rebuilding it
from scratch: external links, VBA, defined names, hidden sheets, cross-sheet
dependency graph, volatile functions, broken references, data connections.

Runs in streaming mode (openpyxl read-only) so multi-hundred-MB workbooks are
handled in bounded memory. Sheet metadata (validations, conditional formats,
merges, protection, freeze panes, declared dimension) is read directly from the
sheet XML in chunks.

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
import posixpath
import re
import sys
import time
import warnings
import zipfile
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path
from urllib.parse import unquote
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


NS = {
    "m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}

VOLATILE = ("OFFSET(", "INDIRECT(", "NOW(", "TODAY(", "RAND(", "RANDBETWEEN(", "CELL(", "INFO(")
LOOKUPS = ("VLOOKUP(", "HLOOKUP(", "XLOOKUP(", "INDEX(", "MATCH(", "SUMIFS(", "SUMIF(", "COUNTIFS(", "SUMPRODUCT(", "GETPIVOTDATA(")
ERROR_VALUES = ("#REF!", "#N/A", "#VALUE!", "#DIV/0!", "#NAME?", "#NUM!", "#NULL!")
BLOAT_MIN_ROWS = 10000  # declared rows above this, and > 2x the real used rows, is flagged as bloat

# 'Sheet Name'!A1  or  SheetName!A1 ; skip [n]Sheet!A1 (external) and #REF!
SHEET_REF_RE = re.compile(r"(?<![\]#A-Za-z0-9_\.])(?:'((?:[^']|'')+)'|([A-Za-z0-9_\.]+))!")
EXT_INDEX_RE = re.compile(r"\[(\d+)\]")
EXT_FILE_RE = re.compile(r"\[[^\]]*\.(xls[xmb]?|xla[m]?|xlt[xm]?|csv)\]", re.I)   # [Book.xlsx] style external path
VBA_PROC_RE = re.compile(r"^\s*(?:Public\s+|Private\s+|Friend\s+)?(?:Static\s+)?(Sub|Function|Property\s+(?:Get|Let|Set))\s+([A-Za-z_][A-Za-z0-9_]*)", re.I | re.M)
VBA_RISK_RE = re.compile(r"\b(Shell|Kill|CreateObject|GetObject|Application\.OnTime|SendKeys|Environ|FileCopy|RmDir|MkDir|Workbooks\.Open|ActiveWorkbook\.SaveAs|DisplayAlerts\s*=\s*False|On Error Resume Next|Sheets\([^)]*\)\.Delete|\.Delete)\b", re.I)

# byte patterns counted in sheet XML (trailing space excludes the plural container elements)
SHEET_XML_COUNTERS = {
    "data_validations": (b"<dataValidation ", b"<x14:dataValidation "),
    "conditional_formats": (b"<cfRule ", b"<x14:cfRule "),
    "merged_ranges": (b"<mergeCell ",),
    "hyperlinks": (b"<hyperlink ",),
    "protected": (b"<sheetProtection",),
    "freeze_panes": (b'state="frozen"',),
    "legacy_drawing": (b"<legacyDrawing",),
    "controls": (b"<control ",),
}
DIMENSION_RE = re.compile(rb'<dimension ref="([A-Z]+\d+(?::[A-Z]+\d+)?)"')


def human_size(n: float) -> str:
    for unit in ("B", "KB", "MB", "GB"):
        if n < 1024:
            return f"{n:.0f} {unit}" if unit == "B" else f"{n:.1f} {unit}"
        n /= 1024
    return f"{n:.1f} TB"


# ----------------------------------------------------------------------------- package-level XML

def workbook_xml(zf: zipfile.ZipFile) -> tuple[ET.Element | None, dict[str, str]]:
    try:
        wb_xml = ET.fromstring(zf.read("xl/workbook.xml"))
        wb_rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    except KeyError:
        return None, {}
    rid_to_target = {r.get("Id"): r.get("Target") for r in wb_rels.findall("rel:Relationship", NS)}
    return wb_xml, rid_to_target


def parse_sheets(wb_xml: ET.Element | None, rid_to_target: dict[str, str]) -> list[dict]:
    """Ordered sheets with name, state and part path, straight from workbook.xml."""
    out = []
    if wb_xml is None:
        return out
    sheets = wb_xml.find("m:sheets", NS)
    if sheets is None:
        return out
    for i, s in enumerate(sheets.findall("m:sheet", NS)):
        rid = s.get(f"{{{NS['r']}}}id")
        target = rid_to_target.get(rid, "") or ""
        part = target.lstrip("/")
        if not part.startswith("xl/"):
            part = "xl/" + part
        out.append({"index": i, "name": s.get("name"), "state": s.get("state", "visible"), "part": part, "sheet_id": s.get("sheetId")})
    return out


def parse_defined_names(wb_xml: ET.Element | None, sheets: list[dict]) -> list[dict]:
    out = []
    if wb_xml is None:
        return out
    dn = wb_xml.find("m:definedNames", NS)
    if dn is None:
        return out
    for d in dn.findall("m:definedName", NS):
        name = d.get("name", "")
        ref = (d.text or "").strip()
        lsid = d.get("localSheetId")
        scope = "workbook"
        if lsid is not None:
            try:
                scope = sheets[int(lsid)]["name"]
            except (ValueError, IndexError):
                scope = f"sheet#{lsid}"
        out.append({
            "name": name,
            "scope": scope,
            "refers_to": ref,
            "broken": "#REF!" in ref,
            # a workbook index [1] or a file name in brackets; plain brackets are structured references (tblPlant[Date])
            "external": bool(EXT_INDEX_RE.search(ref)) or bool(EXT_FILE_RE.search(ref)),
            "hidden": d.get("hidden") == "1",
            "builtin": name.startswith("_xlnm."),
        })
    return out


def parse_external_links(zf: zipfile.ZipFile, wb_xml: ET.Element | None, rid_to_target: dict[str, str]) -> list[dict]:
    """External link targets in workbook order (index as used in formulas: [1], [2] ...)."""
    links = []
    if wb_xml is None:
        return links
    ext_refs = wb_xml.find("m:externalReferences", NS)
    if ext_refs is None:
        return links
    for idx, er in enumerate(ext_refs.findall("m:externalReference", NS), start=1):
        rid = er.get(f"{{{NS['r']}}}id")
        part = rid_to_target.get(rid, "") or ""
        part_path = ("xl/" + part) if not part.startswith("/") else part.lstrip("/")
        rels_path = part_path.replace("externalLinks/", "externalLinks/_rels/") + ".rels"
        target, kind, sheet_names = "", "unknown", []
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
            if link_xml.find("m:ddeLink", NS) is not None:
                kind = "ddeLink"
            if link_xml.find("m:oleLink", NS) is not None:
                kind = "oleLink"
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
    if os.path.isabs(t) or re.match(r"^[A-Za-z]:", t) or t.startswith("\\\\"):
        cand = Path(t)
    else:
        cand = base_dir / t
    try:
        return {"resolved": cand.exists(), "path": str(cand)}
    except OSError:
        return {"resolved": False, "path": str(cand)}


def parse_connections(zf: zipfile.ZipFile) -> list[dict]:
    out = []
    if "xl/connections.xml" in zf.namelist():
        try:
            root = ET.fromstring(zf.read("xl/connections.xml"))
            for c in root.findall("m:connection", NS):
                entry = {"name": c.get("name"), "type": c.get("type"), "description": c.get("description", "")}
                db = c.find("m:dbPr", NS)
                if db is not None:
                    entry["connection"] = (db.get("connection", "") or "")[:300]
                    entry["command"] = (db.get("command", "") or "")[:300]
                out.append(entry)
        except ET.ParseError:
            pass
    return out


def parse_power_query(zf: zipfile.ZipFile) -> list[str]:
    names = []
    for n in zf.namelist():
        if n.startswith("customXml/item") and n.endswith(".xml"):
            try:
                if b"DataMashup" in zf.read(n):
                    names.append(n)
            except KeyError:
                pass
    return names


def parse_table_names(zf: zipfile.ZipFile) -> list[dict]:
    out = []
    for n in zf.namelist():
        if re.match(r"xl/tables/table\d+\.xml$", n):
            try:
                t = ET.fromstring(zf.read(n))
                out.append({"name": t.get("name") or t.get("displayName"), "ref": t.get("ref"), "part": n})
            except ET.ParseError:
                pass
    return out


def count_parts(zf: zipfile.ZipFile, prefix: str) -> int:
    return sum(1 for n in zf.namelist() if n.startswith(prefix))


def sheet_rel_counts(zf: zipfile.ZipFile) -> dict[str, dict]:
    """sheet part -> counts of charts / images / pivot tables / tables / comments via rels."""
    out: dict[str, dict] = {}
    names = set(zf.namelist())
    for n in names:
        m = re.match(r"xl/worksheets/_rels/(sheet\d+\.xml)\.rels$", n)
        if not m:
            continue
        counts = {"charts": 0, "images": 0, "pivot_tables": 0, "tables": 0, "comments": 0}
        try:
            rels = ET.fromstring(zf.read(n))
        except ET.ParseError:
            continue
        for r in rels.findall("rel:Relationship", NS):
            typ = r.get("Type", "").rsplit("/", 1)[-1]
            target = r.get("Target", "")
            if typ == "drawing":
                # relative to xl/worksheets/ (../drawings/x.xml) or absolute in the package (/xl/drawings/x.xml, as openpyxl writes)
                dpath = target.lstrip("/") if target.startswith("/") else posixpath.normpath("xl/worksheets/" + target)
                drels = posixpath.join(posixpath.dirname(dpath), "_rels", posixpath.basename(dpath) + ".rels")
                if drels in names:
                    try:
                        for rr in ET.fromstring(zf.read(drels)).findall("rel:Relationship", NS):
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
        out["xl/worksheets/" + m.group(1)] = counts
    return out


def scan_sheet_xml(zf: zipfile.ZipFile, part: str, chunk: int = 4 * 1024 * 1024) -> dict:
    """Chunked byte scan of a sheet part: declared dimension and feature counts, bounded memory."""
    res = {k: 0 for k in SHEET_XML_COUNTERS}
    res["declared_dimension"] = ""
    res["xml_bytes"] = 0
    try:
        info = zf.getinfo(part)
    except KeyError:
        return res
    res["xml_bytes"] = info.file_size
    overlap = 64
    tail = b""
    first = True
    with zf.open(part) as fh:
        while True:
            data = fh.read(chunk)
            if not data:
                break
            buf = tail + data
            if first:
                m = DIMENSION_RE.search(buf)
                if m:
                    res["declared_dimension"] = m.group(1).decode()
                first = False
            for key, pats in SHEET_XML_COUNTERS.items():
                for p in pats:
                    res[key] += buf.count(p)
            # subtract matches that will be counted again from the overlap region
            tail = data[-overlap:] if len(data) >= overlap else data
            if len(buf) > len(data):  # remove matches fully inside the previous tail, which were counted last round
                prev_tail = buf[: len(buf) - len(data)]
                for key, pats in SHEET_XML_COUNTERS.items():
                    for p in pats:
                        res[key] -= prev_tail.count(p)
    res["protected"] = res["protected"] > 0
    res["freeze_panes"] = res["freeze_panes"] > 0
    return res


def dim_rows_cols(dim: str) -> tuple[int, int]:
    if not dim:
        return 0, 0
    last = dim.split(":")[-1]
    m = re.match(r"([A-Z]+)(\d+)", last)
    if not m:
        return 0, 0
    col = 0
    for ch in m.group(1):
        col = col * 26 + (ord(ch) - 64)
    return int(m.group(2)), col


# ----------------------------------------------------------------------------- VBA

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
            module = {
                "module": vba_filename,
                "lines": len(lines),
                "procedures": [f"{k} {n}" for k, n in procs],
                "sheets_referenced": sorted(set(re.findall(r"(?:Sheets|Worksheets)\(\"([^\"]+)\"\)", code))),
                "workbooks_referenced": sorted(set(re.findall(r"Workbooks(?:\.Open)?\(\"([^\"]+)\"\)", code))),
                "hardcoded_paths": sorted(set(re.findall(r"\"([A-Za-z]:\\[^\"]+|\\\\[^\"]+)\"", code))),
                "risk_keywords": sorted(set(m.group(1) for m in VBA_RISK_RE.finditer(code))),
            }
            result["modules"].append(module)
            result["procedures"] += len(procs)
            result["lines"] += len(lines)
            for _k, n in procs:
                nl = n.lower()
                if nl in ("workbook_open", "auto_open", "workbook_beforeclose", "workbook_beforesave", "auto_close") or nl.startswith("worksheet_"):
                    result["auto_exec"].append(f"{vba_filename}:{n}")
            result["risk_hits"].extend(f"{vba_filename}:{r}" for r in module["risk_keywords"])
    finally:
        vp.close()
    return result


# ----------------------------------------------------------------------------- cell scans

def scan_formula(f: str, own_sheet: str, stats: dict, deps: Counter, ext_use: Counter):
    fu = f.upper()
    stats["formulas"] += 1
    if any(v in fu for v in VOLATILE):
        stats["volatile"] += 1
    if any(l in fu for l in LOOKUPS):
        stats["lookups"] += 1
    if "#REF!" in fu:
        stats["ref_errors_in_formula"] += 1
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
            continue
        if name != own_sheet:
            deps[(own_sheet, name)] += 1


def iter_cells(ws):
    """Stream rows of a read-only sheet without padding to the declared dimension."""
    try:
        ws.reset_dimensions()
    except Exception:
        pass
    if hasattr(ws, "_cells_by_row"):
        return ws._cells_by_row(1, 1, None, None, values_only=False)
    return ws.iter_rows()


def formula_text(v) -> str | None:
    if isinstance(v, str):
        return v if v.startswith("=") else None
    txt = getattr(v, "text", None)  # ArrayFormula
    if txt is not None:
        return txt if str(txt).startswith("=") else "=" + str(txt)
    if v.__class__.__name__ in ("DataTableFormula",):
        return "=TABLE()"
    return None


def inspect_workbook(path: Path) -> dict:
    t0 = time.time()
    info: dict = {
        "file": str(path),
        "name": path.name,
        "size_bytes": path.stat().st_size,
        "size": human_size(path.stat().st_size),
        "modified": datetime.fromtimestamp(path.stat().st_mtime).strftime("%Y-%m-%d %H:%M"),
        "errors": [],
    }
    if not zipfile.is_zipfile(path):
        info["errors"].append("Not an OOXML zip (legacy .xls or .xlsb are not supported by this tool; save as .xlsm)")
        return info

    with zipfile.ZipFile(path) as zf:
        names = zf.namelist()
        wb_xml, rid_to_target = workbook_xml(zf)
        sheet_meta = parse_sheets(wb_xml, rid_to_target)
        info["has_vba_project"] = "xl/vbaProject.bin" in names
        info["external_links"] = parse_external_links(zf, wb_xml, rid_to_target)
        for l in info["external_links"]:
            l.update(resolve_link_target(l["target"], path.parent))
        info["defined_names"] = parse_defined_names(wb_xml, sheet_meta)
        info["connections"] = parse_connections(zf)
        info["power_query_parts"] = parse_power_query(zf)
        info["tables"] = parse_table_names(zf)
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
            "activex": count_parts(zf, "xl/activeX/"),
            "embeddings": count_parts(zf, "xl/embeddings/"),
        }
        info["uncompressed_bytes"] = sum(i.file_size for i in zf.infolist())
        largest = sorted(zf.infolist(), key=lambda i: i.file_size, reverse=True)[:8]
        info["largest_parts"] = [{"part": i.filename, "size": human_size(i.file_size)} for i in largest]
        rel_counts = sheet_rel_counts(zf)
        xml_meta = {s["part"]: scan_sheet_xml(zf, s["part"]) for s in sheet_meta}

    info["vba"] = analyse_vba(path) if info["has_vba_project"] else {"present": False, "modules": [], "procedures": 0, "lines": 0, "auto_exec": [], "risk_hits": [], "error": None}

    # ---- formula pass (streaming)
    by_name = {s["name"]: s for s in sheet_meta}
    sheets: list[dict] = []
    deps: Counter = Counter()
    ext_use: Counter = Counter()
    total: Counter = Counter()
    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=False, keep_links=True)
    except Exception as exc:
        info["errors"].append(f"openpyxl could not open workbook: {exc}")
        return info
    try:
        for ws in wb.worksheets:
            meta = by_name.get(ws.title, {"state": getattr(ws, "sheet_state", "visible"), "part": ""})
            xm = xml_meta.get(meta.get("part"), {})
            st = {"formulas": 0, "volatile": 0, "lookups": 0, "ref_errors_in_formula": 0, "array_formulas": 0, "external_link_formulas": 0}
            values = 0
            last_row = 0
            last_col = 0
            samples: list[str] = []
            try:
                for row in iter_cells(ws):
                    for c in row:
                        v = getattr(c, "value", None)
                        if v is None:
                            continue
                        values += 1
                        r_, c_ = getattr(c, "row", 0), getattr(c, "column", 0)
                        if r_ > last_row:
                            last_row = r_
                        if c_ > last_col:
                            last_col = c_
                        f = formula_text(v)
                        if f is None:
                            continue
                        if not isinstance(v, str):
                            st["array_formulas"] += 1
                        scan_formula(f, ws.title, st, deps, ext_use)
                        if len(samples) < 5 and ("[" in f or "!" in f):
                            samples.append(f"{get_column_letter(c_)}{r_}: {f[:120]}")
            except Exception as exc:
                info["errors"].append(f"cell scan failed on '{ws.title}': {exc}")
            decl_rows, decl_cols = dim_rows_cols(xm.get("declared_dimension", ""))
            used = f"A1:{get_column_letter(last_col)}{last_row}" if last_row and last_col else "(empty)"
            bloated = decl_rows >= BLOAT_MIN_ROWS and decl_rows > 2 * max(last_row, 1)
            entry = {
                "sheet": ws.title,
                "state": meta.get("state", "visible"),
                "declared_dimension": xm.get("declared_dimension", ""),
                "used_range": used,
                "last_row": last_row,
                "last_col": last_col,
                "bloated_dimension": bloated,
                "xml_bytes": xm.get("xml_bytes", 0),
                "cells_with_values": values,
                **st,
                "merged_ranges": xm.get("merged_ranges", 0),
                "data_validations": xm.get("data_validations", 0),
                "conditional_formats": xm.get("conditional_formats", 0),
                "hyperlinks": xm.get("hyperlinks", 0),
                "protected": xm.get("protected", False),
                "freeze_panes": xm.get("freeze_panes", False),
                "form_controls": xm.get("controls", 0) + xm.get("legacy_drawing", 0),
                **rel_counts.get(meta.get("part", ""), {"charts": 0, "images": 0, "pivot_tables": 0, "tables": 0, "comments": 0}),
                "link_formula_samples": samples,
            }
            sheets.append(entry)
            for k, v in st.items():
                total[k] += v
            total["cells_with_values"] += values
        for cs in getattr(wb, "chartsheets", []):
            sheets.append({"sheet": cs.title, "state": by_name.get(cs.title, {}).get("state", "visible"), "type": "chartsheet"})
    finally:
        wb.close()

    info["sheets"] = sheets
    info["totals"] = dict(total)
    info["cross_sheet_dependencies"] = [{"from": a, "to": b, "formulas": n} for (a, b), n in sorted(deps.items(), key=lambda kv: -kv[1])]
    known = {s["sheet"] for s in sheets}
    info["dangling_sheet_refs"] = sorted({b for (a, b) in deps if b not in known})
    for l in info["external_links"]:
        l["formulas_using"] = ext_use.get(l["index"], 0)

    # ---- cached value error pass (streaming)
    err: Counter = Counter()
    err_by_sheet: dict[str, Counter] = defaultdict(Counter)
    try:
        wbv = openpyxl.load_workbook(path, read_only=True, data_only=True, keep_links=False)
        try:
            for ws in wbv.worksheets:
                for row in iter_cells(ws):
                    for c in row:
                        v = getattr(c, "value", None)
                        if isinstance(v, str) and v in ERROR_VALUES:
                            err[v] += 1
                            err_by_sheet[ws.title][v] += 1
        finally:
            wbv.close()
    except Exception as exc:
        info["errors"].append(f"cached value scan failed: {exc}")
    info["cached_errors"] = dict(err)
    info["cached_errors_by_sheet"] = {k: dict(v) for k, v in err_by_sheet.items()}
    info["scan_seconds"] = round(time.time() - t0, 1)
    return info


# ----------------------------------------------------------------------------- reporting

def md_table(headers: list[str], rows: list[list]) -> str:
    out = ["| " + " | ".join(headers) + " |", "|" + "|".join("---" for _ in headers) + "|"]
    for r in rows:
        out.append("| " + " | ".join(str(x).replace("|", "\\|") if x is not None else "" for x in r) + " |")
    return "\n".join(out)


def render_report(results: list[dict], base: Path) -> str:
    L: list[str] = []
    L.append(f"# Excel workbook inspection\n\nSource: `{base}`  \nGenerated: {datetime.now():%Y-%m-%d %H:%M}\n")

    rows = []
    for r in results:
        if "sheets" not in r:
            rows.append([r["name"], r["size"], "ERROR: " + "; ".join(r.get("errors", [])), "", "", "", "", "", ""])
            continue
        hidden = sum(1 for s in r["sheets"] if s.get("state") != "visible")
        dead = sum(1 for l in r["external_links"] if not l.get("resolved"))
        vba = r["vba"]
        rows.append([
            r["name"], r["size"], len(r["sheets"]), hidden,
            f"{len(r['external_links'])} ({dead} unresolved)",
            r["totals"].get("formulas", 0),
            r["totals"].get("external_link_formulas", 0),
            f"{vba['procedures']} procs / {vba['lines']} lines" if vba["present"] else ("bin present, unparsed" if r.get("has_vba_project") else "none"),
            sum(r["cached_errors"].values()),
        ])
    L.append("## Summary\n")
    L.append(md_table(["Workbook", "Size", "Sheets", "Hidden", "External links", "Formulas", "Ext-link formulas", "VBA", "Cached errors"], rows))
    L.append("")

    L.append("## Cross-workbook link map\n")
    names = {r["name"].lower() for r in results}
    graph_rows = []
    for r in results:
        for l in r.get("external_links", []):
            tgt = unquote(l["target"])
            tgt_name = tgt.replace("\\", "/").rsplit("/", 1)[-1]
            graph_rows.append([r["name"], f"[{l['index']}]", l.get("type", ""), tgt[:110], "ok" if l.get("resolved") else "MISSING", "yes" if tgt_name.lower() in names else "no", l.get("formulas_using", 0), ", ".join(l["sheets_referenced"][:6])])
    L.append(md_table(["From", "Idx", "Type", "Target", "On disk", "In this set", "Formulas", "Sheets referenced"], graph_rows) if graph_rows else "_No external links found._")
    L.append("")

    for r in results:
        L.append(f"\n---\n\n## {r['name']}\n")
        L.append(f"- Size {r['size']} on disk, {human_size(r.get('uncompressed_bytes', 0))} uncompressed. Modified {r['modified']}. Scan {r.get('scan_seconds', '?')} s.")
        for e in r.get("errors", []):
            L.append(f"- **Error:** {e}")
        if "sheets" not in r:
            continue
        pc = r["part_counts"]
        L.append(f"- Parts: {pc['worksheets']} worksheets, {pc['charts']} charts, {pc['pivot_tables']} pivot tables ({pc['pivot_caches']} caches), {pc['tables']} tables, {pc['images']} media files, {pc['query_tables']} query tables, {pc['slicers']} slicers, {pc['activex']} ActiveX, {pc['embeddings']} embeddings.")
        L.append(f"- Data connections: {len(r['connections'])}; Power Query parts: {len(r['power_query_parts'])}.")
        L.append("- Largest parts: " + ", ".join(f"{p['part'].split('/')[-1]} {p['size']}" for p in r["largest_parts"][:5]))
        L.append("")

        L.append("### Sheets\n")
        srows = []
        for s in r["sheets"]:
            if s.get("type") == "chartsheet":
                srows.append([s["sheet"], s["state"], "chartsheet", "", "", "", "", "", "", "", "", ""])
                continue
            used = s["used_range"] + (" BLOAT" if s["bloated_dimension"] else "")
            srows.append([
                s["sheet"], s["state"], s["declared_dimension"], used, s["cells_with_values"], s["formulas"],
                s["external_link_formulas"], s["volatile"], s["lookups"], s["data_validations"],
                s["conditional_formats"], f"{s['charts']}c/{s['pivot_tables']}p/{s['images']}i/{s['tables']}t",
            ])
        L.append(md_table(["Sheet", "State", "Declared", "Used", "Values", "Formulas", "Ext-link", "Volatile", "Lookups", "DV", "CF", "Charts/Pivots/Imgs/Tables"], srows))
        bloat = [s["sheet"] for s in r["sheets"] if s.get("bloated_dimension")]
        if bloat:
            L.append(f"\nDeclared range far larger than real content (dead rows, slows the file): {', '.join(bloat)}")
        prot = [s["sheet"] for s in r["sheets"] if s.get("protected")]
        if prot:
            L.append(f"\nProtected sheets: {', '.join(prot)}")
        L.append("")

        L.append("### Cross-sheet dependencies (formula count, top 40)\n")
        deps = r["cross_sheet_dependencies"][:40]
        L.append(md_table(["From sheet", "Reads from", "Formulas"], [[d["from"], d["to"], d["formulas"]] for d in deps]) if deps else "_None._")
        if r["dangling_sheet_refs"]:
            L.append(f"\n**Dangling sheet references (sheet no longer exists):** {', '.join(r['dangling_sheet_refs'])}")
        L.append("")

        dn = r["defined_names"]
        broken = [d for d in dn if d["broken"]]
        ext = [d for d in dn if d["external"] and not d["broken"]]
        L.append(f"### Defined names: {len(dn)} total, {len(broken)} broken (#REF!), {len(ext)} pointing to external workbooks, {sum(1 for d in dn if d['hidden'])} hidden\n")
        show = broken + ext
        if show:
            L.append(md_table(["Name", "Scope", "Refers to", "Status"], [[d["name"], d["scope"], d["refers_to"][:100], "BROKEN" if d["broken"] else "external"] for d in show[:60]]))
        L.append("")

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

        if r["connections"]:
            L.append("### Data connections\n")
            L.append(md_table(["Name", "Type", "Connection", "Command"], [[c.get("name"), c.get("type"), c.get("connection", "")[:100], c.get("command", "")[:100]] for c in r["connections"]]))
            L.append("")

        if r["tables"]:
            L.append("### Excel tables (ListObjects)\n")
            L.append(", ".join(f"{t['name']} ({t['ref']})" for t in r["tables"][:40]))
            L.append("")

        if r["cached_errors"]:
            L.append("### Cached error values\n")
            L.append(md_table(["Sheet"] + list(ERROR_VALUES), [[sh] + [cnt.get(e, 0) for e in ERROR_VALUES] for sh, cnt in sorted(r["cached_errors_by_sheet"].items(), key=lambda kv: -sum(kv[1].values()))[:30]]))
            L.append("")

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
        print(f"inspecting {f.name} ({human_size(f.stat().st_size)}) ...", flush=True)
        try:
            r = inspect_workbook(f)
        except Exception as exc:
            r = {"file": str(f), "name": f.name, "size": human_size(f.stat().st_size), "size_bytes": f.stat().st_size, "modified": "", "errors": [f"unhandled: {exc!r}"]}
        results.append(r)
        (out / (f.stem + ".json")).write_text(json.dumps(r, indent=2, default=str), encoding="utf-8")
        print(f"  done in {r.get('scan_seconds', '?')} s", flush=True)

    base = files[0].parent if len({f.parent for f in files}) == 1 else Path(os.path.commonpath([str(f) for f in files]))
    (out / "inspection_report.md").write_text(render_report(results, base), encoding="utf-8")
    print(f"\nreport written to {out / 'inspection_report.md'}")
    if not HAVE_OLEVBA:
        print("note: install oletools for VBA analysis:  pip install oletools")


if __name__ == "__main__":
    main()
'@
Write-TextFile (Join-Path $Root "01_Tools\xl_inspect.py") $content_tools_xl_inspect_py $false
$content_tools_PythonEnv_ps1 = @'
<#
.SYNOPSIS
    Python discovery and installation helpers shared by the Mako Production System scripts.
.DESCRIPTION
    Dot-source this file from a script in the same folder:

        . (Join-Path $here "PythonEnv.ps1")

    Functions:
        Find-Python       path of the first real Python 3.8+ on PATH or behind the py launcher,
                          otherwise the newest one in the usual install folders (per-user,
                          Program Files, C:\Python3x, Anaconda and Miniconda), or $null. Every
                          candidate is probed for its version; the Microsoft Store stub under
                          \Microsoft\WindowsApps\ fails that probe (exit code 9009) and is
                          skipped, while a real Store Python at the same path is accepted.
        Install-Python    installs Python 3.12 for the current user from python.org, silently
                          (InstallAllUsers=0 InstallLauncherAllUsers=0 PrependPath=1, no admin
                          rights, through the Windows proxy when one is configured), with winget
                          as fallback, then refreshes PATH for the running session.
        Test-PyModule     $true when the given interpreter can import the module.
        Ensure-PyModules  pip installs each missing module (passing the Windows proxy when one
                          applies), retrying with --user, and throws with a manual command when
                          a module still cannot be imported.
        Resolve-Python    Find-Python, then Install-Python unless -NoInstall, then Find-Python
                          again. Returns the interpreter path or throws.

    These functions run native executables (python, pip, winget). Windows PowerShell 5.1 turns
    native stderr output into terminating errors when $ErrorActionPreference is Stop, so each
    function sets Continue for its own scope and checks $LASTEXITCODE explicitly. Callers should
    do the same for the native commands they run themselves.
.NOTES
    Windows PowerShell 5.1 and PowerShell 7 compatible. Save as UTF-8 with BOM, CRLF.
#>

function Test-RealPython {
    param([string]$Exe)
    $ErrorActionPreference = "Continue"
    if ([string]::IsNullOrWhiteSpace($Exe)) { return $false }
    if (-not (Test-Path -LiteralPath $Exe)) { return $false }
    # The Microsoft Store stub (python.exe under \Microsoft\WindowsApps\) prints a message and exits 9009 when
    # given arguments, so the exit code test below rejects it; a real Store Python at the same path passes.
    $v = $null
    try { $v = & $Exe -c "import sys; print(sys.version_info[0]*100 + sys.version_info[1])" 2>$null }
    catch { return $false }
    if ($LASTEXITCODE -ne 0) { return $false }
    $n = 0
    if (-not [int]::TryParse(("$v").Trim(), [ref]$n)) { return $false }
    return ($n -ge 308)
}

function Find-Python {
    $ErrorActionPreference = "Continue"
    $candidates = New-Object System.Collections.Generic.List[string]
    foreach ($name in @("python.exe", "python3.exe")) {
        Get-Command $name -All -ErrorAction SilentlyContinue | ForEach-Object { $candidates.Add($_.Source) }
    }
    $pyl = Get-Command py.exe -ErrorAction SilentlyContinue
    if ($pyl) {
        $resolved = $null
        try { $resolved = & $pyl.Source -3 -c "import sys; print(sys.executable)" 2>$null } catch { }
        if ($LASTEXITCODE -eq 0 -and $resolved) { $candidates.Add(("$resolved").Trim()) }
    }
    $globs = @(
        "$env:LOCALAPPDATA\Programs\Python\Python3*\python.exe",
        "$env:ProgramFiles\Python3*\python.exe",
        "${env:ProgramFiles(x86)}\Python3*\python.exe",
        "C:\Python3*\python.exe",
        "$env:LOCALAPPDATA\anaconda3\python.exe",
        "$env:USERPROFILE\anaconda3\python.exe",
        "$env:USERPROFILE\miniconda3\python.exe",
        "$env:ProgramData\anaconda3\python.exe"
    )
    foreach ($g in $globs) {
        # newest first by the file version resource, not by name (a name sort puts Python39 above Python312)
        Get-ChildItem -Path $g -ErrorAction SilentlyContinue | Sort-Object { $_.VersionInfo.FileVersionRaw } -Descending | ForEach-Object { $candidates.Add($_.FullName) }
    }
    foreach ($c in $candidates) { if (Test-RealPython $c) { return $c } }
    return $null
}

function Use-SystemProxy {
    # Windows PowerShell 5.1 uses the proxy from Internet Options but sends it no credentials, so an
    # authenticating site proxy answers 407. Attach the current user's credentials to the default proxy.
    try {
        $proxy = [System.Net.WebRequest]::GetSystemWebProxy()
        $proxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
        [System.Net.WebRequest]::DefaultWebProxy = $proxy
    } catch { }
}

function Get-PipProxyArgs {
    # pip reads a static proxy from Internet Options but not a proxy script (PAC). Resolve the proxy
    # Windows would use for PyPI and return it as pip arguments, or an empty array for a direct connection.
    try {
        $target = [uri]"https://pypi.org/"
        $p = [System.Net.WebRequest]::GetSystemWebProxy().GetProxy($target)
        if ($p -and $p.Host -ne $target.Host) { return @("--proxy", $p.AbsoluteUri) }
    } catch { }
    return @()
}

function Install-Python {
    param([string]$Version = "3.12.10")   # last 3.12 release with a Windows installer
    $ErrorActionPreference = "Continue"
    $ProgressPreference = "SilentlyContinue"   # the 5.1 progress bar makes Invoke-WebRequest many times slower
    $arch = if ($env:PROCESSOR_ARCHITECTURE -eq "ARM64") { "arm64" } else { "amd64" }
    $url = "https://www.python.org/ftp/python/$Version/python-$Version-$arch.exe"
    $installer = Join-Path $env:TEMP "python-$Version-$arch.exe"
    Write-Host "Python 3 not found. Installing Python $Version for the current user (no admin rights, 25 MB download, about 100 MB installed, 1-2 minutes) ..."
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Use-SystemProxy
        Invoke-WebRequest -Uri $url -OutFile $installer -UseBasicParsing -ErrorAction Stop
        # InstallLauncherAllUsers=0: the py launcher defaults to a per-machine component, which asks for
        # elevation even in a per-user /quiet install.
        $p = Start-Process -FilePath $installer -ArgumentList "/quiet InstallAllUsers=0 InstallLauncherAllUsers=0 PrependPath=1 Include_launcher=1 Include_test=0 Include_doc=0" -Wait -PassThru
        # 0 = ok, 3010 = ok but a restart is pending
        if ($p.ExitCode -ne 0 -and $p.ExitCode -ne 3010) { throw "installer exit code $($p.ExitCode)" }
        Write-Host "Python installed."
    } catch {
        Write-Warning "python.org install failed: $($_.Exception.Message). Trying winget ..."
        $wg = Get-Command winget.exe -ErrorAction SilentlyContinue
        if (-not $wg) { throw "Neither python.org nor winget is available. Install Python 3 manually from https://www.python.org/downloads/windows/ (tick 'Add python.exe to PATH'), then re-run this script." }
        & $wg.Source install --id Python.Python.3.12 -e --scope user --silent --accept-package-agreements --accept-source-agreements 2>&1 | ForEach-Object { Write-Host "  $_" }
        if ($LASTEXITCODE -ne 0) { throw "winget install failed with exit code $LASTEXITCODE. Install Python 3 manually from https://www.python.org/downloads/windows/ and re-run." }
    }
    # refresh PATH for this session so the new interpreter is visible without opening a new window;
    # Machine then User as in a fresh window, and the entries this process already had are kept
    $env:Path = [Environment]::GetEnvironmentVariable("Path", "Machine") + ";" + [Environment]::GetEnvironmentVariable("Path", "User") + ";" + $env:Path
}

function Test-PyModule {
    param([string]$Python, [string]$Module)
    $ErrorActionPreference = "Continue"
    if ([string]::IsNullOrWhiteSpace($Python)) { return $false }
    # find_spec writes nothing to stdout or stderr; Out-Null keeps the pipeline clean regardless
    & $Python -c "import importlib.util, sys; sys.exit(0 if importlib.util.find_spec('$Module') else 1)" 2>$null | Out-Null
    return ($LASTEXITCODE -eq 0)
}

function Ensure-PyModules {
    param([string]$Python, [string[]]$Modules)
    $ErrorActionPreference = "Continue"
    $proxyArgs = @(Get-PipProxyArgs)
    foreach ($m in $Modules) {
        if (Test-PyModule $Python $m) { continue }
        Write-Host "Installing Python package $m ..."
        if ($proxyArgs.Count -gt 0) { Write-Host "  via proxy $($proxyArgs[1])" }
        & $Python -m pip install --quiet --disable-pip-version-check @proxyArgs $m 2>&1 | ForEach-Object { Write-Host "  $_" }
        if (-not (Test-PyModule $Python $m)) {
            & $Python -m pip install --quiet --disable-pip-version-check --user @proxyArgs $m 2>&1 | ForEach-Object { Write-Host "  $_" }
        }
        if (-not (Test-PyModule $Python $m)) {
            $manual = (@("`"$Python`"", "-m", "pip", "install") + $proxyArgs + @($m)) -join " "
            throw "Could not install Python package '$m'. Install it manually:  $manual"
        }
    }
}

function Resolve-Python {
    param([switch]$NoInstall)
    $ErrorActionPreference = "Continue"
    $py = Find-Python
    if ($py) { return $py }
    if ($NoInstall) { throw "Python 3 not found and -NoInstall was given. Install Python 3 from https://www.python.org/downloads/windows/ (tick 'Add python.exe to PATH') and re-run." }
    Install-Python | Out-Null
    $py = Find-Python
    if (-not $py) { throw "Python was installed but could not be located. Open a new PowerShell window and re-run this script." }
    return $py
}
'@
Write-TextFile (Join-Path $Root "01_Tools\PythonEnv.ps1") $content_tools_PythonEnv_ps1 $true
$content_tools_Invoke_XlInspect_ps1 = @'
<#
.SYNOPSIS
    Runs xl_inspect.py against the PMC operation report workbooks.
.DESCRIPTION
    Finds a real Python 3 (ignoring the Microsoft Store stub). If none exists,
    installs Python 3.12 for the current user from python.org (no admin rights),
    falling back to winget. Installs openpyxl and oletools if missing, then writes
    inspection_report.md and one JSON per workbook to the output folder.
    Python discovery and installation live in PythonEnv.ps1 next to this script.
.PARAMETER Source
    Folder holding the workbooks. Defaults to <root>\02_Source\2026 when it exists,
    otherwise the 2026 operation report folder on X:.
.PARAMETER Out
    Output folder for the report. Defaults to <root>\03_Inspection in the C:\MakoPS
    layout, otherwise <Source>\_inspection.
.PARAMETER NoInstall
    Never install Python; fail with instructions instead.
.EXAMPLE
    .\Invoke-XlInspect.ps1
    .\Invoke-XlInspect.ps1 -Source "X:\13_Management\01_GM_Dashboard\01_Operation_Report\01_Excel\2026"
.NOTES
    Mapped drives such as X: are not visible in an elevated (Run as administrator)
    PowerShell session. Run this from a normal PowerShell window.
#>
[CmdletBinding()]
param(
    [string]$Source = "",
    [string]$Out = "",
    [switch]$NoInstall
)

# Continue, not Stop: Windows PowerShell 5.1 turns native stderr output into terminating
# errors under Stop, and pip / openpyxl legitimately write warnings to stderr.
$ErrorActionPreference = "Continue"
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$script = Join-Path $here "xl_inspect.py"
$root = Split-Path -Parent $here

$envFile = Join-Path $here "PythonEnv.ps1"
if (-not (Test-Path -LiteralPath $envFile)) { throw "PythonEnv.ps1 not found next to this script: $envFile" }
. $envFile
if (-not (Test-Path -LiteralPath $script)) { throw "xl_inspect.py not found next to this script: $script" }

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
# Absolute paths without a trailing backslash: Windows PowerShell 5.1 hands 'D:\Ops Reports\' to python.exe
# as "D:\Ops Reports\" and the C runtime reads \" as a literal quote, which swallows the next argument.
$Source = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Source)
$Out = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Out)
if ($Source -notmatch '^[A-Za-z]:\\$') { $Source = $Source.TrimEnd('\') }
if ($Out -notmatch '^[A-Za-z]:\\$') { $Out = $Out.TrimEnd('\') }

# ---------------------------------------------------------------- python and packages
$py = Resolve-Python -NoInstall:$NoInstall
Write-Host "Using Python: $py"
Ensure-PyModules $py @("openpyxl", "oletools")

# ---------------------------------------------------------------- run
Write-Host "Inspecting workbooks in $Source"
& $py $script $Source -o $Out
if ($LASTEXITCODE -ne 0) { throw "xl_inspect.py failed with exit code $LASTEXITCODE" }

$report = Join-Path $Out "inspection_report.md"
Write-Host ""
Write-Host "Report: $report"
Write-Host "Attach the whole '$Out' folder (report + JSON files) to the session."
'@
Write-TextFile (Join-Path $Root "01_Tools\Invoke-XlInspect.ps1") $content_tools_Invoke_XlInspect_ps1 $true
$content_tools_Build_MPS_ps1 = @'
<#
.SYNOPSIS
    Builds the Mako Production System workbooks (blank and DEMO) from the schema.
.DESCRIPTION
    Finds or installs Python 3 (PythonEnv.ps1), makes sure openpyxl is available, then runs
    mps\build_workbook.py twice:

        <Root>\04_MPS\MPS_<Year>.xlsx        blank workbook for production use
        <Root>\04_MPS\MPS_<Year>_DEMO.xlsx   the same workbook with <SampleDays> days of synthetic DEMO data

    Existing files are overwritten, so close them in Excel before re-running. The generator is
    looked up next to this script first (<Root>\01_Tools\mps\build_workbook.py in the C:\MakoPS
    layout, <repo>\mps\build_workbook.py when this script runs from the repository), then under
    <Root>\01_Tools\mps; the path used is printed.
.PARAMETER Root
    Working folder root. Default C:\MakoPS. Output goes to <Root>\04_MPS.
.PARAMETER Year
    Reporting year of the workbook. Default 2026.
.PARAMETER SampleDays
    Days of synthetic data in the DEMO workbook. Default 74. 0 skips the DEMO workbook.
.PARAMETER NoInstall
    Never install Python; fail with instructions instead.
.PARAMETER NoOpen
    Do not open the output folder in Explorer when done.
.EXAMPLE
    .\Build-MPS.ps1
    .\Build-MPS.ps1 -Root D:\MakoPS -Year 2027 -SampleDays 30
#>
[CmdletBinding()]
param(
    [string]$Root = "C:\MakoPS",
    [int]$Year = 2026,
    [int]$SampleDays = 74,
    [switch]$NoInstall,
    [switch]$NoOpen
)

# Continue, not Stop: Windows PowerShell 5.1 turns native stderr output into terminating
# errors under Stop, and pip / openpyxl legitimately write warnings to stderr.
$ErrorActionPreference = "Continue"
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
# Absolute root: the .NET file test below resolves relative paths against the process directory, not the PowerShell location
$Root = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Root)

$envFile = Join-Path $here "PythonEnv.ps1"
if (-not (Test-Path -LiteralPath $envFile)) { throw "PythonEnv.ps1 not found next to this script: $envFile" }
. $envFile

# ---------------------------------------------------------------- locate the generator
# The copies next to this script win over a deployed copy under <Root>, so a run from the repository
# builds with the repository schema even when C:\MakoPS exists.
$candidates = @(
    (Join-Path $here "mps\build_workbook.py"),
    (Join-Path (Split-Path -Parent $here) "mps\build_workbook.py"),
    (Join-Path $Root "01_Tools\mps\build_workbook.py")
)
$generator = $null
foreach ($c in $candidates) { if (Test-Path -LiteralPath $c) { $generator = $c; break } }
if (-not $generator) { throw ("build_workbook.py not found. Looked in:`n  " + ($candidates -join "`n  ")) }
$schema = Join-Path (Split-Path -Parent $generator) "schema\mps_schema.json"
if (-not (Test-Path -LiteralPath $schema)) { throw "Schema not found next to the generator: $schema" }
Write-Host "Using generator: $generator"
Write-Host "Using schema:    $schema"

$outDir = Join-Path $Root "04_MPS"
if (-not (Test-Path -LiteralPath $outDir)) { New-Item -ItemType Directory -Force -Path $outDir | Out-Null }

# ---------------------------------------------------------------- python and packages
$py = Resolve-Python -NoInstall:$NoInstall
Write-Host "Using Python: $py"
Ensure-PyModules $py @("openpyxl")

# ---------------------------------------------------------------- build
$blank = Join-Path $outDir ("MPS_{0}.xlsx" -f $Year)
$demo  = Join-Path $outDir ("MPS_{0}_DEMO.xlsx" -f $Year)

function Invoke-Generator([string[]]$Arguments, [string]$Target) {
    if (Test-Path -LiteralPath $Target) {
        # a workbook that is open in Excel cannot be overwritten; fail before spending time on the build
        try { $fs = [System.IO.File]::Open($Target, "Open", "ReadWrite", "None"); $fs.Close() }
        catch { throw "Cannot overwrite $Target. Close it in Excel and re-run." }
    }
    Write-Host ("Building {0} ..." -f (Split-Path -Leaf $Target))
    & $py $generator @Arguments
    if ($LASTEXITCODE -ne 0) { throw "build_workbook.py failed with exit code $LASTEXITCODE" }
    if (-not (Test-Path -LiteralPath $Target)) { throw "build_workbook.py finished but $Target was not written" }
}

Invoke-Generator @("--out", $blank, "--year", "$Year") $blank
if ($SampleDays -gt 0) {
    Invoke-Generator @("--out", $demo, "--year", "$Year", "--sample-days", "$SampleDays") $demo
}

# ---------------------------------------------------------------- report
Write-Host ""
Write-Host "MPS workbooks written to $outDir"
Write-Host ("  {0}  ({1:N1} MB)  blank, for production use" -f $blank, ((Get-Item -LiteralPath $blank).Length / 1MB))
if ($SampleDays -gt 0) {
    Write-Host ("  {0}  ({1:N1} MB)  {2} days of synthetic DEMO data" -f $demo, ((Get-Item -LiteralPath $demo).Length / 1MB), $SampleDays)
}
Write-Host ""
Write-Host "Open the DEMO workbook first to see the Dashboard, Daily_Report and Monthly_Summary populated."
Write-Host "Enter real data in the blank workbook only. Read the README sheet inside the workbook before the first entry."
if (-not $NoOpen) { try { Start-Process explorer.exe $outDir } catch { } }
'@
Write-TextFile (Join-Path $Root "01_Tools\Build-MPS.ps1") $content_tools_Build_MPS_ps1 $true
$content_tools_Export_DailyReport_ps1 = @'
<#
.SYNOPSIS
    Exports the Daily_Report sheet of an MPS workbook to PDF with Excel, optionally into an Outlook mail.
.DESCRIPTION
    Uses Excel COM automation (the workbook contains no macros). The script opens the workbook
    read-only, writes the report date into Daily_Report!C5, lets Excel calculate, exports the
    Daily_Report sheet to

        <OutDir>\Mako_Daily_Report_yyyy-MM-dd.pdf

    and closes Excel without saving, so the workbook is left exactly as it was. With -Email an
    Outlook message with the PDF attached is displayed for review; nothing is sent automatically.
    Requires Microsoft Excel (and classic Outlook for -Email) on this PC.
.PARAMETER Workbook
    Path to the MPS workbook. Default: <Root>\04_MPS\MPS_<year of Date>.xlsx.
.PARAMETER Date
    Report date, for example 2026-03-15. Default: yesterday.
.PARAMETER OutDir
    Folder for the PDF. Default: <workbook folder>\Reports (created when missing).
.PARAMETER Root
    Working folder root used for the default workbook path. Default C:\MakoPS.
.PARAMETER Sheet
    Report sheet name. Default Daily_Report.
.PARAMETER DateCell
    Cell holding the report date on that sheet. Default C5.
.PARAMETER Email
    Create an Outlook message with the PDF attached and display it. The user reviews and sends it.
.PARAMETER To
    Recipients for -Email. One or more addresses.
.PARAMETER Cc
    Copy recipients for -Email.
.PARAMETER Subject
    Mail subject. Default "Mako Daily Production Report <date>".
.PARAMETER Open
    Open the PDF when done.
.EXAMPLE
    .\Export-DailyReport.ps1
    .\Export-DailyReport.ps1 -Workbook C:\MakoPS\04_MPS\MPS_2026_DEMO.xlsx -Date 2026-03-15 -Open
    .\Export-DailyReport.ps1 -Date 2026-03-15 -Email -To gm@example.com,ops@example.com
#>
[CmdletBinding()]
param(
    [string]$Workbook = "",
    [datetime]$Date = (Get-Date).Date.AddDays(-1),
    [string]$OutDir = "",
    [string]$Root = "C:\MakoPS",
    [string]$Sheet = "Daily_Report",
    [string]$DateCell = "C5",
    [switch]$Email,
    [string[]]$To = @(),
    [string[]]$Cc = @(),
    [string]$Subject = "",
    [switch]$Open
)

# No native commands run here, so Stop is safe and makes every COM failure land in the catch/finally blocks.
$ErrorActionPreference = "Stop"

# Excel COM binds through the thread culture. A Windows regional format with no matching Office
# language (for example French (Senegal) with English Office) raises 0x80028018, "Old format or
# invalid type library", on the first property set. en-US also keeps English day and month names
# in the file name and the mail subject.
$culture = [System.Globalization.CultureInfo]::GetCultureInfo("en-US")
[System.Threading.Thread]::CurrentThread.CurrentCulture = $culture
[System.Threading.Thread]::CurrentThread.CurrentUICulture = $culture

$Date = $Date.Date
if ([string]::IsNullOrWhiteSpace($Workbook)) { $Workbook = Join-Path $Root ("04_MPS\MPS_{0}.xlsx" -f $Date.Year) }
if (-not (Test-Path -LiteralPath $Workbook)) { throw "Workbook not found: $Workbook  (run Build-MPS.ps1 first, or pass -Workbook)" }
$Workbook = (Resolve-Path -LiteralPath $Workbook).ProviderPath
if ([string]::IsNullOrWhiteSpace($OutDir)) { $OutDir = Join-Path (Split-Path -Parent $Workbook) "Reports" }
if (-not (Test-Path -LiteralPath $OutDir)) { New-Item -ItemType Directory -Force -Path $OutDir | Out-Null }
$OutDir = (Resolve-Path -LiteralPath $OutDir).ProviderPath
$pdf = Join-Path $OutDir ("Mako_Daily_Report_{0:yyyy-MM-dd}.pdf" -f $Date)
if ($Email -and $To.Count -eq 0) { throw "-Email needs at least one recipient, for example: -To name@example.com" }

function Remove-ComRef($obj) {
    # Releases the runtime callable wrapper so the Excel / Outlook process can exit.
    if ($null -ne $obj) {
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($obj) } catch { }
    }
}

# ---------------------------------------------------------------- Excel: set the date, calculate, export
$excel = $null; $books = $null; $wb = $null; $sheets = $null; $ws = $null; $cell = $null
try {
    try { $excel = New-Object -ComObject Excel.Application }
    catch {
        throw ("Microsoft Excel is not installed on this computer, or COM automation is blocked, so the PDF cannot be produced here. " +
               "Run this script on a PC with Excel, or open the workbook, set {0}!{1} to the report date and use File > Export > Create PDF." -f $Sheet, $DateCell)
    }
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false
    $excel.AskToUpdateLinks = $false
    $excel.EnableEvents = $false
    try { $excel.AutomationSecurity = 3 } catch { }   # msoAutomationSecurityForceDisable: the workbook has no macros

    Write-Host "Opening $Workbook (read-only)"
    $books = $excel.Workbooks
    $wb = $books.Open($Workbook, 0, $true)            # UpdateLinks = 0, ReadOnly = true
    $sheets = $wb.Worksheets
    try { $ws = $sheets.Item($Sheet) } catch { throw "Sheet '$Sheet' not found in $Workbook" }

    # C5 is unlocked in the protected sheet, so it can be written without unprotecting anything.
    $cell = $ws.Range($DateCell)
    $cell.Value2 = $Date.ToOADate()

    $excel.CalculateFull()
    $deadline = (Get-Date).AddMinutes(5)
    while ($excel.CalculationState -ne 0 -and (Get-Date) -lt $deadline) { Start-Sleep -Milliseconds 250 }   # 0 = xlDone

    # Warn when the requested date lies beyond the entered data (the report would be blank).
    $nm = $null; $rng = $null
    try {
        $nm = $wb.Names.Item("cfg_LastDataDate")
        $rng = $nm.RefersToRange
        $last = $rng.Value2
        if ($last -is [double] -and $last -gt 0) {
            $lastDate = [datetime]::FromOADate($last)
            if ($Date -gt $lastDate) {
                Write-Warning ("Last entered data is {0:yyyy-MM-dd}; the report for {1:yyyy-MM-dd} will show no figures." -f $lastDate, $Date)
            }
        }
    } catch { }
    finally { Remove-ComRef $rng; Remove-ComRef $nm }

    if (Test-Path -LiteralPath $pdf) { Remove-Item -LiteralPath $pdf -Force }
    Write-Host ("Exporting {0} for {1:ddd dd-MMM-yyyy} to PDF" -f $Sheet, $Date)
    # ExportAsFixedFormat(Type, Filename, Quality, IncludeDocProperties, IgnorePrintAreas): xlTypePDF = 0, xlQualityStandard = 0
    $ws.ExportAsFixedFormat(0, $pdf, 0, $true, $false)
    if (-not (Test-Path -LiteralPath $pdf)) { throw "Excel did not write $pdf" }
    Write-Host "PDF: $pdf"
}
finally {
    if ($null -ne $wb) { try { $wb.Close($false) } catch { } }          # never save
    if ($null -ne $excel) { try { $excel.Quit() } catch { } }
    foreach ($o in @($cell, $ws, $sheets, $wb, $books, $excel)) { Remove-ComRef $o }
    $cell = $null; $ws = $null; $sheets = $null; $wb = $null; $books = $null; $excel = $null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect()
}

# ---------------------------------------------------------------- Outlook: draft the mail (displayed, not sent)
if ($Email) {
    $ol = $null; $mail = $null; $att = $null
    try {
        try { $ol = New-Object -ComObject Outlook.Application }
        catch { throw "Classic Microsoft Outlook is not available on this computer (the new Outlook has no COM automation). The PDF is at $pdf; attach it to a mail manually." }
        if ([string]::IsNullOrWhiteSpace($Subject)) { $Subject = ("Mako Daily Production Report {0:ddd dd-MMM-yyyy}" -f $Date) }
        $mail = $ol.CreateItem(0)                     # olMailItem
        $mail.To = ($To -join "; ")
        if ($Cc.Count -gt 0) { $mail.CC = ($Cc -join "; ") }
        $mail.Subject = $Subject
        $mail.Body = ("Please find attached the Mako daily production report for {0:dddd dd MMMM yyyy}.`r`n`r`nSource workbook: {1}`r`n" -f $Date, (Split-Path -Leaf $Workbook))
        $att = $mail.Attachments
        [void]$att.Add($pdf)
        $mail.Display()
        Write-Host "Outlook message opened for review. Check the recipients and send it yourself."
    }
    finally {
        foreach ($o in @($att, $mail, $ol)) { Remove-ComRef $o }
        $att = $null; $mail = $null; $ol = $null
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
}

if ($Open) { Start-Process -FilePath $pdf }
'@
Write-TextFile (Join-Path $Root "01_Tools\Export-DailyReport.ps1") $content_tools_Export_DailyReport_ps1 $true
$content_tools_README_md = @'
# tools

Scripts that run on the site PCs (Windows PowerShell 5.1, no admin rights) and the
build script that packs them into the single `Setup-MakoPS.ps1` bootstrap.

## Folder layout on the PC

`Setup-MakoPS.ps1` creates and fills this layout (default root `C:\MakoPS`):

| Folder | Content |
|---|---|
| `01_Tools` | the scripts below, `xl_inspect.py`, `README.md` (this file) |
| `01_Tools\mps` | MPS generator: `build_workbook.py`, `build_ddl.py`, `validate_workbook.py`, `schema\mps_schema.json` |
| `02_Source\2026` | working copies of the PMC operation workbooks (originals on X: are never touched) |
| `03_Inspection` | `inspection_report.md` plus one JSON per inspected workbook |
| `04_MPS` | `MPS_<year>.xlsx` (blank, for production use) and `MPS_<year>_DEMO.xlsx` (synthetic data) |

Run everything from a normal PowerShell window. Mapped drives such as `X:` are not visible in
an elevated (Run as administrator) session, so the copy step would find nothing there. If the
execution policy blocks scripts, start them with
`powershell -ExecutionPolicy Bypass -File <script>`.

## PythonEnv.ps1

Shared helper, dot-sourced by the other scripts (`. (Join-Path $here "PythonEnv.ps1")`). No
parameters of its own.

| Function | Purpose |
|---|---|
| `Find-Python` | first real Python 3.8+ on PATH or behind the `py` launcher, otherwise the newest one in the per-user and Program Files install folders, Anaconda and Miniconda. Every candidate is probed for its version, which rejects the Microsoft Store stub (exit code 9009) but accepts a real Store Python. |
| `Install-Python` | silent per-user install of Python 3.12 from python.org (`InstallAllUsers=0 InstallLauncherAllUsers=0 PrependPath=1`, no admin rights, through the Windows proxy with the user's credentials), winget as fallback, then refreshes PATH for the running session |
| `Test-PyModule` | `$true` when the interpreter can import a module |
| `Ensure-PyModules` | `pip install` of each missing module (with `--proxy` when Windows routes PyPI through a proxy), retried with `--user`, with a manual command in the error when that fails too |
| `Resolve-Python` | `Find-Python`, then `Install-Python` unless `-NoInstall`, then `Find-Python` again |

The functions run native executables, so they set `$ErrorActionPreference = "Continue"` in
their own scope and test `$LASTEXITCODE` explicitly: Windows PowerShell 5.1 turns anything a
native command writes to stderr into a terminating error under `Stop`, and pip does write
warnings there.

## Invoke-XlInspect.ps1

Runs `xl_inspect.py` over a folder of workbooks.

```powershell
.\Invoke-XlInspect.ps1                       # <root>\02_Source\2026 to <root>\03_Inspection
.\Invoke-XlInspect.ps1 -Source "X:\13_Management\01_GM_Dashboard\01_Operation_Report\01_Excel\2026"
.\Invoke-XlInspect.ps1 -NoInstall            # fail instead of installing Python
```

Needs `openpyxl` and `oletools`; both are installed when missing.

## Build-MPS.ps1

Builds the Mako Production System workbooks from `mps\schema\mps_schema.json`.

```powershell
.\Build-MPS.ps1                              # C:\MakoPS\04_MPS\MPS_2026.xlsx and MPS_2026_DEMO.xlsx
.\Build-MPS.ps1 -Root D:\MakoPS -Year 2027 -SampleDays 30
.\Build-MPS.ps1 -NoOpen                      # do not open 04_MPS in Explorer
```

| Parameter | Default | Meaning |
|---|---|---|
| `-Root` | `C:\MakoPS` | layout root; output goes to `<Root>\04_MPS` |
| `-Year` | 2026 | reporting year of the workbook |
| `-SampleDays` | 74 | days of synthetic data in the DEMO workbook, 0 skips it |
| `-NoInstall` | | never install Python |
| `-NoOpen` | | do not open the output folder |

The generator is taken from `<Root>\01_Tools\mps\build_workbook.py`; when the script runs from
this repository, `<repo>\mps\build_workbook.py` is used. Each workbook takes a few seconds.
Existing files are overwritten, so close them in Excel first (the script checks and stops with
a message if a file is locked). Open the DEMO workbook to see the Dashboard, Daily_Report and
Monthly_Summary populated; enter real data only in the blank workbook.

## Export-DailyReport.ps1

Produces the daily PDF from an MPS workbook with Excel COM automation, so the workbook itself
needs no macros. The script opens the workbook read-only, writes the date into
`Daily_Report!C5`, lets Excel calculate, exports the sheet with `ExportAsFixedFormat` and closes
Excel without saving. Requires Microsoft Excel on the PC; `-Email` additionally requires classic
Outlook (the new Outlook has no COM automation).

```powershell
.\Export-DailyReport.ps1                                   # yesterday, C:\MakoPS\04_MPS\MPS_<year>.xlsx
.\Export-DailyReport.ps1 -Workbook C:\MakoPS\04_MPS\MPS_2026_DEMO.xlsx -Date 2026-03-15 -Open
.\Export-DailyReport.ps1 -Date 2026-03-15 -Email -To gm@example.com,ops@example.com
```

| Parameter | Default | Meaning |
|---|---|---|
| `-Workbook` | `<Root>\04_MPS\MPS_<year of Date>.xlsx` | workbook to export from |
| `-Date` | yesterday | report date, write it as `yyyy-MM-dd` |
| `-OutDir` | `<workbook folder>\Reports` | where the PDF goes, created when missing |
| `-Root` | `C:\MakoPS` | only used for the default workbook path |
| `-Sheet`, `-DateCell` | `Daily_Report`, `C5` | report sheet and its date cell |
| `-Email -To a,b -Cc c -Subject s` | | display an Outlook mail with the PDF attached (nothing is sent) |
| `-Open` | | open the PDF when done |

Output: `<OutDir>\Mako_Daily_Report_yyyy-MM-dd.pdf`. A warning is printed when the requested
date lies after `cfg_LastDataDate`, because the report would then be blank. To run it every
morning, create a Task Scheduler task with the action
`powershell -ExecutionPolicy Bypass -File C:\MakoPS\01_Tools\Export-DailyReport.ps1`.

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

The scan streams every sheet (openpyxl read-only) and reads sheet features straight from the
XML, so memory stays flat on large workbooks and a bloated declared range (for example
`A1:Z1048576` with 400 real rows) costs nothing. Such sheets are flagged `BLOAT` in the report.

Any platform:

```bash
pip install openpyxl oletools
python tools/xl_inspect.py <folder-or-files> -o <output-dir>
```

Output: `inspection_report.md` plus one `<workbook>.json` per file. It is also the quickest
check of a freshly built MPS workbook: `python tools/xl_inspect.py C:\MakoPS\04_MPS -o C:\MakoPS\03_Inspection\mps`.

## build_setup.py

Packs the scripts above, `xl_inspect.py`, this README and the `mps` generator files
(`build_workbook.py`, `build_ddl.py`, `validate_workbook.py`, `schema\mps_schema.json`,
`README.md`; the optional ones are skipped with a warning when absent) into
`Setup-MakoPS.ps1` at the repository root. Each file travels in a single-quoted PowerShell
here-string, so nothing is interpolated; a file containing a line that starts with `'@` cannot
be embedded and the build stops with an error.

```bash
python tools/build_setup.py            # regenerate Setup-MakoPS.ps1, then verify the round trip
python tools/build_setup.py --check    # only verify the existing Setup-MakoPS.ps1 against the sources
```

The verification extracts every here-string from the generated file and compares it with its
source, the target path and the BOM flag. Re-run the build after changing any embedded file.
`Setup-MakoPS.ps1` accepts `-Root`, `-Source`, `-Year`, `-SkipCopy`, `-SkipRun` and `-SkipBuild`.
It resolves Python once (installing it when missing) and passes `-NoInstall` to the two child
scripts; when Python cannot be resolved the audit and the build are skipped with a warning and
the setup still finishes. The audit and the build run in try/catch blocks, so one failure does
not stop the other, and a single Explorer window is opened on the root at the end.

## Conventions

- PowerShell files: Windows PowerShell 5.1 compatible (no ternary operator, no null-coalescing,
  no PowerShell 7 only cmdlets), saved as UTF-8 with BOM and CRLF.
- Python files: Python 3.8+, `openpyxl` only (plus `oletools` for the VBA scan).
- British English, metric units, gold in troy ounces, grades in g/t.
'@
Write-TextFile (Join-Path $Root "01_Tools\README.md") $content_tools_README_md $false
$content_mps_build_workbook_py = @'
#!/usr/bin/env python3
"""
build_workbook.py - generate the Mako Production System (MPS) Excel workbook
from mps_schema.json and the KPI catalogue below.

Design rules baked in:
  * one workbook, no external links, no VBA, no hidden sheets
  * every input is an Excel table (ListObject) with data validation
  * all reporting reads from Calc_Daily, one row per day, built with SUMIFS
    over the tables; the report and dashboard never touch the inputs directly
  * the same schema drives the SQL DDL for the web phase (build_ddl.py)

Usage:
    python build_workbook.py --out MPS_2026.xlsx [--year 2026] [--sample-days 74]

--sample-days N fills N days of synthetic data (clearly marked DEMO) so the
report can be reviewed before real data exists.
"""
from __future__ import annotations

import argparse
import calendar
import json
import random
import re
import sys
from datetime import date, timedelta
from pathlib import Path

import openpyxl
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.axis import DateAxis
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Protection, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.pagebreak import Break
from openpyxl.worksheet.properties import PageSetupProperties
from openpyxl.worksheet.table import Table, TableFormula, TableStyleInfo

HERE = Path(__file__).resolve().parent
DEFAULT_SCHEMA = HERE / "schema" / "mps_schema.json"

# ----------------------------------------------------------------------------- styling
FONT_NAME = "Calibri"
C_HEAD = "1F3864"        # header fill (dark blue), white text
C_HEAD_CALC = "595959"   # calculated column header (dark grey)
C_LABEL = "F2F2F2"       # label row
C_SECTION = "D9E1F2"     # report section band
C_INPUT_CELL = "FFFFFF"
C_KEY = "FFF2CC"         # report date input
C_WARN = "FFC7CE"        # conditional highlight: missing or unknown value
TAB = {"report": "2F5597", "input": "548235", "config": "7F7F7F", "calc": "3A3A3A"}

FMT = {
    "t": "#,##0", "oz": "#,##0", "g/t": "0.00", "%": "0.0%", "h": "#,##0.0", "kg": "#,##0", "kWh": "#,##0",
    "m3": "#,##0", "L": "#,##0", "m": "#,##0", "bcm": "#,##0", "t/h": "#,##0", "kg/t": "0.00", "kWh/t": "0.0",
    "m3/t": "0.00", "ratio": "0.00", "rate": "0.00", "days": "0", "int": "0", "date": "dd-mmm-yyyy", "text": "@", "number": "#,##0.00",
}
TROY = "cfg_TroyOz"
# schema constants exposed as named Config cells; {NAME} placeholders in table calcs and KPI formulas resolve to the name
CONSTANTS = {"TROY_OZ_G": ("cfg_TroyOz", "Troy ounce (g)"),
             "HOURS_PER_DAY": ("cfg_HoursPerDay", "Hours per day (fleet and mill calendar hours)"),
             "RATE_BASIS_HOURS": ("cfg_RateBasisHours", "Injury rate basis, hours (TRIFR and LTIFR are injuries per this many hours)")}
PLACEHOLDERS = {"TROY": TROY, **{k: v[0] for k, v in CONSTANTS.items()}}
DEFAULT_DATE_FORMULA = "=MAX(cfg_YearStart,MIN(TODAY()-1,cfg_YearEnd,IF(N(cfg_LastDataDate)>0,cfg_LastDataDate,cfg_YearEnd)))"

thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)


def font(bold=False, color="000000", size=10, italic=False):
    return Font(name=FONT_NAME, bold=bold, color=color, size=size, italic=italic)


def fill(hex_):
    return PatternFill("solid", start_color=hex_, end_color=hex_)


def const_subst(expr: str) -> str:
    """Replace {TROY}, {HOURS_PER_DAY}, {RATE_BASIS_HOURS} with their Config names."""
    for k, v in PLACEHOLDERS.items():
        expr = expr.replace("{" + k + "}", v)
    return expr


def num_text(v) -> str:
    """Number as Excel formula text: 5000000 not 5000000.0."""
    return str(int(v)) if float(v).is_integer() else str(v)


def xl_date_text(expr: str) -> str:
    """Locale-independent 'dd Mon yyyy' text for a date expression (TEXT() format letters change with the Excel UI language)."""
    return (f'DAY({expr})&" "&CHOOSE(MONTH({expr}),"Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec")'
            f'&" "&YEAR({expr})')


# ----------------------------------------------------------------------------- KPI catalogue
# kind: sum   -> Day column (formula given), MTD and YTD running totals
#       ratio -> num / den per period, built from other KPI columns
#       custom-> explicit Day formula only (rates, ages)
#       point -> value at the date (not additive)
# budget: name of a tblBudget column; ratio budgets are derived automatically when num/den are budgeted
def K(key, label, unit, kind, **kw):
    d = dict(key=key, label=label, unit=unit, kind=kind)
    d.update(kw)
    return d


def sumifs(tbl, col, *extra):
    parts = [f"{tbl}[{col}]", f"{tbl}[Date]", "{d}"]
    for c, v in extra:
        parts += [f"{tbl}[{c}]", v]
    return "SUMIFS(" + ",".join(parts) + ")"


KPIS = [
    # ---- mining
    K("Ore_Mined_t", "Ore mined (ex-pit)", "t", "sum", day=sumifs("tblMovement", "Tonnes_t", ("Is_Ore", "1"), ("Source_Type", '"Pit"')), budget="Ore_Mined_t"),
    K("Ore_Mined_oz", "Contained gold mined", "oz", "sum", day=sumifs("tblMovement", "Ore_oz", ("Source_Type", '"Pit"')), budget="Ore_oz"),
    # graded ore tonnes carry the grade denominator so an ore row entered before its grade does not dilute the grade;
    # budget alias keeps the derived grade budget (the budget assumes every budgeted tonne is graded)
    K("Ore_Graded_t", "Ore mined with a grade (grade basis)", "t", "sum",
      day=sumifs("tblMovement", "Tonnes_t", ("Is_Ore", "1"), ("Source_Type", '"Pit"'), ("Grade_gpt", '"<>"')), budget="Ore_Mined_t"),
    K("Ore_Grade_gpt", "Ore grade mined", "g/t", "ratio", num="{Ore_Mined_oz}*" + TROY, den="{Ore_Graded_t}"),
    K("Ore_Ungraded_t", "Ore mined without a grade (check)", "t", "sum", day="{Ore_Mined_t}-{Ore_Graded_t}"),
    K("Waste_t", "Waste mined", "t", "sum", day=sumifs("tblMovement", "Tonnes_t", ("Is_Ore", "0"), ("Source_Type", '"Pit"')), budget="Waste_t"),
    K("TMM_t", "Total material moved", "t", "sum", day="{Ore_Mined_t}+{Waste_t}", budget="TMM_t"),
    K("Strip_Ratio", "Strip ratio (waste : ore)", "ratio", "ratio", num="{Waste_t}", den="{Ore_Mined_t}"),
    K("Rehandle_t", "Rehandle from stockpiles", "t", "sum", day=sumifs("tblMovement", "Tonnes_t", ("Source_Type", '"Stockpile"'))),
    K("Moved_Total_t", "Total tonnes moved (all movement rows)", "t", "sum", day=sumifs("tblMovement", "Tonnes_t")),
    K("Unclassified_t", "Movements with unknown source type (check)", "t", "sum", day="{Moved_Total_t}-{TMM_t}-{Rehandle_t}"),
    K("Drill_m", "Drilled", "m", "sum", day=sumifs("tblMiningDaily", "Drill_m"), budget="Drill_m"),
    K("GC_Drill_m", "Grade control drilled", "m", "sum", day=sumifs("tblMiningDaily", "GC_Drill_m")),
    K("Blast_t", "Blasted", "t", "sum", day=sumifs("tblMiningDaily", "Blast_t")),
    K("Explosives_kg", "Explosives", "kg", "sum", day=sumifs("tblMiningDaily", "Explosives_kg")),
    K("Powder_Factor_kgpt", "Powder factor", "kg/t", "ratio", num="{Explosives_kg}", den="{Blast_t}"),
    K("Diesel_Mining_L", "Mining diesel", "L", "sum", day=sumifs("tblMiningDaily", "Diesel_L"), budget="Diesel_L"),
    K("Dewatering_m3", "Pit dewatering", "m3", "sum", day=sumifs("tblMiningDaily", "Dewatering_m3")),
    K("Exc_Calendar_h", "Excavator calendar hours", "h", "sum", day=sumifs("tblFleet", "Calendar_h", ("Equipment_Class", '"Excavator"'))),
    K("Exc_Available_h", "Excavator available hours", "h", "sum", day=sumifs("tblFleet", "Available_h", ("Equipment_Class", '"Excavator"'))),
    K("Exc_Operating_h", "Excavator operating hours", "h", "sum", day=sumifs("tblFleet", "Operating_h", ("Equipment_Class", '"Excavator"'))),
    K("Exc_Availability_pct", "Excavator availability", "%", "ratio", num="{Exc_Available_h}", den="{Exc_Calendar_h}"),
    K("Exc_Utilisation_pct", "Excavator utilisation", "%", "ratio", num="{Exc_Operating_h}", den="{Exc_Available_h}"),
    K("Trk_Calendar_h", "Truck calendar hours", "h", "sum", day=sumifs("tblFleet", "Calendar_h", ("Equipment_Class", '"Haul Truck"'))),
    K("Trk_Available_h", "Truck available hours", "h", "sum", day=sumifs("tblFleet", "Available_h", ("Equipment_Class", '"Haul Truck"'))),
    K("Trk_Operating_h", "Truck operating hours", "h", "sum", day=sumifs("tblFleet", "Operating_h", ("Equipment_Class", '"Haul Truck"'))),
    K("Trk_Availability_pct", "Truck availability", "%", "ratio", num="{Trk_Available_h}", den="{Trk_Calendar_h}"),
    K("Trk_Utilisation_pct", "Truck utilisation", "%", "ratio", num="{Trk_Operating_h}", den="{Trk_Available_h}"),
    K("Exc_Productivity_tph", "Excavator productivity (TMM + rehandle)", "t/h", "ratio", num="{TMM_t}+{Rehandle_t}", den="{Exc_Operating_h}"),
    K("Trk_Productivity_tph", "Truck productivity (TMM + rehandle)", "t/h", "ratio", num="{TMM_t}+{Rehandle_t}", den="{Trk_Operating_h}"),
    # ---- processing
    K("Crushed_t", "Crushed", "t", "sum", day=sumifs("tblPlant", "Crushed_t")),
    K("Milled_t", "Milled", "t", "sum", day=sumifs("tblPlant", "Milled_t"), budget="Milled_t"),
    K("Feed_oz", "Contained gold in feed", "oz", "sum", day=sumifs("tblPlant", "Feed_oz"), budget="Feed_oz"),
    K("Tails_oz", "Gold lost to tails", "oz", "sum", day=sumifs("tblPlant", "Tails_oz")),
    K("Recovered_oz", "Gold recovered", "oz", "sum", day=sumifs("tblPlant", "Recovered_oz"), budget="Recovered_oz"),
    # period grades and recovery are paired with the tonnes that carry the assay, so a day whose head or tails grade
    # is still outstanding is left out of the ratio instead of counting as zero grade or 100 percent recovery;
    # budget aliases keep the derived budgets (the budget has no blank days)
    K("Milled_Graded_t", "Milled with a head grade (head grade basis)", "t", "sum", day=sumifs("tblPlant", "Milled_t", ("Head_Grade_gpt", '"<>"')), budget="Milled_t"),
    K("Milled_Tails_Graded_t", "Milled with a tails grade (tails grade basis)", "t", "sum", day=sumifs("tblPlant", "Milled_t", ("Tails_Grade_gpt", '"<>"'))),
    K("Feed_Recon_oz", "Contained gold in feed on days with a tails grade (recovery basis)", "oz", "sum", day=sumifs("tblPlant", "Feed_oz", ("Tails_Grade_gpt", '"<>"')), budget="Feed_oz"),
    K("Milled_Timed_t", "Milled on days with run hours (throughput basis)", "t", "sum", day=sumifs("tblPlant", "Milled_t", ("Mill_Run_h", '"<>"'))),
    K("Head_Grade_gpt", "Head grade", "g/t", "ratio", num="{Feed_oz}*" + TROY, den="{Milled_Graded_t}"),
    K("Tails_Grade_gpt", "Tails grade", "g/t", "ratio", num="{Tails_oz}*" + TROY, den="{Milled_Tails_Graded_t}"),
    K("Recovery_pct", "Recovery", "%", "ratio", num="{Recovered_oz}", den="{Feed_Recon_oz}"),
    K("Gravity_Gold_oz", "Gravity gold", "oz", "sum", day=sumifs("tblPlant", "Gravity_Gold_oz")),
    K("Gravity_Share_pct", "Gravity share of gold recovered", "%", "ratio", num="{Gravity_Gold_oz}", den="{Recovered_oz}"),
    K("Gold_Poured_oz", "Gold poured", "oz", "sum", day=sumifs("tblPlant", "Gold_Poured_oz"), budget="Gold_Poured_oz"),
    K("Mill_Run_h", "Mill run hours", "h", "sum", day=sumifs("tblPlant", "Mill_Run_h")),
    K("Mill_Planned_Maint_h", "Mill planned maintenance", "h", "sum", day=sumifs("tblPlant", "Mill_Planned_Maint_h")),
    K("Mill_Unplanned_Down_h", "Mill unplanned downtime", "h", "sum", day=sumifs("tblPlant", "Mill_Unplanned_Down_h")),
    K("Mill_Calendar_h", "Mill calendar hours (days with any mill hours entered)", "h", "sum",
      day='IF(COUNTIFS(tblPlant[Date],{d},tblPlant[Mill_Run_h],"<>")+COUNTIFS(tblPlant[Date],{d},tblPlant[Mill_Planned_Maint_h],"<>")'
          '+COUNTIFS(tblPlant[Date],{d},tblPlant[Mill_Unplanned_Down_h],"<>")>0,{HOURS_PER_DAY},0)'),
    K("Crusher_Run_h", "Crusher run hours", "h", "sum", day=sumifs("tblPlant", "Crusher_Run_h")),
    K("Throughput_tph", "Mill throughput", "t/h", "ratio", num="{Milled_Timed_t}", den="{Mill_Run_h}"),
    K("Mill_Availability_pct", "Mill availability", "%", "ratio", num="{Mill_Calendar_h}-{Mill_Planned_Maint_h}-{Mill_Unplanned_Down_h}", den="{Mill_Calendar_h}"),
    K("Mill_Utilisation_pct", "Mill utilisation", "%", "ratio", num="{Mill_Run_h}", den="{Mill_Calendar_h}-{Mill_Planned_Maint_h}-{Mill_Unplanned_Down_h}"),
    K("Cyanide_kg", "Cyanide", "kg", "sum", day=sumifs("tblPlant", "Cyanide_kg")),
    K("Lime_kg", "Lime", "kg", "sum", day=sumifs("tblPlant", "Lime_kg")),
    K("Grinding_Media_kg", "Grinding media", "kg", "sum", day=sumifs("tblPlant", "Grinding_Media_kg")),
    K("Power_kWh", "Plant power", "kWh", "sum", day=sumifs("tblPlant", "Power_kWh")),
    K("Raw_Water_m3", "Raw water", "m3", "sum", day=sumifs("tblPlant", "Raw_Water_m3")),
    K("Cyanide_kgpt", "Cyanide consumption", "kg/t", "ratio", num="{Cyanide_kg}", den="{Milled_t}"),
    K("Lime_kgpt", "Lime consumption", "kg/t", "ratio", num="{Lime_kg}", den="{Milled_t}"),
    K("Media_kgpt", "Grinding media consumption", "kg/t", "ratio", num="{Grinding_Media_kg}", den="{Milled_t}"),
    K("Power_kWhpt", "Power consumption", "kWh/t", "ratio", num="{Power_kWh}", den="{Milled_t}"),
    K("Water_m3pt", "Raw water consumption", "m3/t", "ratio", num="{Raw_Water_m3}", den="{Milled_t}"),
    K("GIC_oz", "Gold in circuit (end of day)", "oz", "point", day="IF(SUMIFS(tblPlant[GIC_oz],tblPlant[Date],{d})>0,SUMIFS(tblPlant[GIC_oz],tblPlant[Date],{d}),\"\")"),
    # ---- gold
    K("Gold_Shipped_oz", "Gold shipped", "oz", "sum", day=sumifs("tblGold", "Gold_oz", ("Type", '"Shipment"'))),
    K("Gold_Sold_oz", "Gold sold", "oz", "sum", day=sumifs("tblGold", "Gold_oz", ("Type", '"Sale"'))),
    K("Dore_Shipped_kg", "Dore shipped", "kg", "sum", day=sumifs("tblGold", "Dore_kg", ("Type", '"Shipment"'))),
    # ---- safety
    K("Hours_Worked", "Hours worked", "h", "sum", day=sumifs("tblSafety", "Hours_Total"), budget="Hours_Worked"),
    K("Fatality", "Fatalities", "int", "sum", day=sumifs("tblSafety", "Fatality")),
    K("LTI", "Lost time injuries", "int", "sum", day=sumifs("tblSafety", "LTI")),
    K("RWI", "Restricted work injuries", "int", "sum", day=sumifs("tblSafety", "RWI")),
    K("MTI", "Medical treatment injuries", "int", "sum", day=sumifs("tblSafety", "MTI")),
    K("FAI", "First aid injuries", "int", "sum", day=sumifs("tblSafety", "FAI")),
    K("Recordables", "Recordable injuries", "int", "sum", day=sumifs("tblSafety", "Recordables")),
    K("Near_Miss", "Near misses", "int", "sum", day=sumifs("tblSafety", "Near_Miss")),
    K("Hazard_Reports", "Hazard reports", "int", "sum", day=sumifs("tblSafety", "Hazard_Reports")),
    K("HPI", "High potential incidents", "int", "sum", day=sumifs("tblSafety", "HPI")),
    K("Env_Incidents", "Environmental incidents", "int", "sum", day=sumifs("tblSafety", "Env_Incidents")),
    K("Community_Incidents", "Community incidents", "int", "sum", day=sumifs("tblSafety", "Community_Incidents")),
    K("Vehicle_Incidents", "Vehicle incidents", "int", "sum", day=sumifs("tblSafety", "Vehicle_Incidents")),
    K("Toolbox_Talks", "Toolbox talks", "int", "sum", day=sumifs("tblSafety", "Toolbox_Talks")),
    K("Inspections", "Inspections and audits", "int", "sum", day=sumifs("tblSafety", "Inspections")),
    # rolling window: daily rows for the last 12 months plus month to date; history months fill any month in the
    # window before the current month that has no daily hours (prior year, or this year before go-live)
    K("TRIFR_12m", "TRIFR, rolling 12 months + MTD", "rate", "custom",
      day='IFERROR((SUMIFS(tblSafety[Recordables],tblSafety[Date],">="&EDATE({ms},-12),tblSafety[Date],"<="&{d})'
          '+SUMIFS(tblSafetyHistory[Recordables],tblSafetyHistory[Month],">="&EDATE({ms},-12),tblSafetyHistory[Month],"<"&{ms},tblSafetyHistory[Use],1))'
          '/(SUMIFS(tblSafety[Hours_Total],tblSafety[Date],">="&EDATE({ms},-12),tblSafety[Date],"<="&{d})'
          '+SUMIFS(tblSafetyHistory[Hours_Total],tblSafetyHistory[Month],">="&EDATE({ms},-12),tblSafetyHistory[Month],"<"&{ms},tblSafetyHistory[Use],1))*{RATE_BASIS_HOURS},"")'),
    K("LTIFR_12m", "LTIFR, rolling 12 months + MTD", "rate", "custom",
      day='IFERROR((SUMIFS(tblSafety[LTI],tblSafety[Date],">="&EDATE({ms},-12),tblSafety[Date],"<="&{d})'
          '+SUMIFS(tblSafetyHistory[LTI],tblSafetyHistory[Month],">="&EDATE({ms},-12),tblSafetyHistory[Month],"<"&{ms},tblSafetyHistory[Use],1))'
          '/(SUMIFS(tblSafety[Hours_Total],tblSafety[Date],">="&EDATE({ms},-12),tblSafety[Date],"<="&{d})'
          '+SUMIFS(tblSafetyHistory[Hours_Total],tblSafetyHistory[Month],">="&EDATE({ms},-12),tblSafetyHistory[Month],"<"&{ms},tblSafetyHistory[Use],1))*{RATE_BASIS_HOURS},"")'),
    K("Days_Since_LTI", "Days since last LTI", "days", "custom",
      day='IF(MAX(N(cfg_PriorLTIDate),_xlfn.MAXIFS(tblSafety[Date],tblSafety[LTI],">0",tblSafety[Date],"<="&{d}))=0,"",'
          '{d}-MAX(N(cfg_PriorLTIDate),_xlfn.MAXIFS(tblSafety[Date],tblSafety[LTI],">0",tblSafety[Date],"<="&{d})))'),
]
KPI = {k["key"]: k for k in KPIS}

# Daily report layout: (section, [(kpi, show_budget)])
REPORT_LAYOUT = [
    ("SAFETY", [("Hours_Worked", True), ("Recordables", False), ("LTI", False), ("RWI", False), ("MTI", False), ("FAI", False),
                ("Near_Miss", False), ("Hazard_Reports", False), ("HPI", False), ("Env_Incidents", False), ("Community_Incidents", False),
                ("Vehicle_Incidents", False), ("TRIFR_12m", False), ("LTIFR_12m", False), ("Days_Since_LTI", False)]),
    ("MINING", [("Ore_Mined_t", True), ("Ore_Grade_gpt", True), ("Ore_Mined_oz", True), ("Ore_Ungraded_t", False), ("Waste_t", True), ("TMM_t", True),
                ("Strip_Ratio", True), ("Rehandle_t", False), ("Unclassified_t", False), ("Drill_m", True), ("GC_Drill_m", False), ("Blast_t", False), ("Powder_Factor_kgpt", False), ("Diesel_Mining_L", True),
                ("Exc_Availability_pct", False), ("Exc_Utilisation_pct", False), ("Exc_Productivity_tph", False),
                ("Trk_Availability_pct", False), ("Trk_Utilisation_pct", False), ("Trk_Productivity_tph", False)]),
    ("PROCESSING", [("Crushed_t", False), ("Milled_t", True), ("Head_Grade_gpt", True), ("Feed_oz", True), ("Recovery_pct", True), ("Recovered_oz", True),
                    ("Gravity_Gold_oz", False), ("Gravity_Share_pct", False), ("Gold_Poured_oz", True), ("Throughput_tph", False), ("Mill_Run_h", False), ("Mill_Availability_pct", False),
                    ("Mill_Utilisation_pct", False), ("Cyanide_kgpt", False), ("Lime_kgpt", False), ("Power_kWhpt", False), ("GIC_oz", False)]),
    ("GOLD", [("Gold_Shipped_oz", False), ("Gold_Sold_oz", False), ("Dore_Shipped_kg", False)]),
]
MONTHLY_KPIS = ["Hours_Worked", "Recordables", "LTI", "TRIFR_12m", "Ore_Mined_t", "Ore_Grade_gpt", "Ore_Mined_oz", "Waste_t", "TMM_t", "Strip_Ratio",
                "Drill_m", "Milled_t", "Head_Grade_gpt", "Recovery_pct", "Recovered_oz", "Gold_Poured_oz", "Gold_Sold_oz", "Throughput_tph",
                "Mill_Availability_pct", "Mill_Utilisation_pct", "Cyanide_kgpt", "Power_kWhpt"]
PAGE_BREAK_BEFORE = "PROCESSING"   # Daily_Report prints on two portrait pages
COMMENT_LINES = 15
COMMENT_ROW_PT = 36               # three lines of 9 pt text per comment; rows stay resizable on the protected sheet


# ----------------------------------------------------------------------------- builder
class Builder:
    def __init__(self, schema: dict, year: int, sample_days: int = 0, seed: int = 7, report_date: date | None = None):
        self.s = schema
        self.year = year
        self.sample_days = sample_days
        self.report_date = report_date
        self.rng = random.Random(seed)
        self.wb = openpyxl.Workbook()
        self.wb.remove(self.wb.active)
        self.days = [date(year, 1, 1) + timedelta(days=i) for i in range((date(year, 12, 31) - date(year, 1, 1)).days + 1)]
        self.tables: dict[str, dict] = {}     # table name -> {sheet, first_row, last_row, cols{name: letter}}
        self.cd_cols: dict[str, str] = {}     # Calc_Daily header -> column letter
        self.sample = {}                      # synthetic data (for the validator)
        self.formula_count = 0

    # ---- helpers -----------------------------------------------------------------
    def name(self, nm: str, ref: str):
        self.wb.defined_names[nm] = DefinedName(nm, attr_text=ref)

    def dyn_range(self, sheet: str, col_letter: str, first_row: int) -> str:
        """Dynamic range from the first data row to the last non-empty cell; counts from the first data row only,
        so the title, label and header cells above the list do not shorten it."""
        c = col_letter
        return f"{sheet}!${c}${first_row}:INDEX({sheet}!${c}:${c},{first_row - 1}+COUNTA({sheet}!${c}${first_row}:${c}$1048576))"

    def set_widths(self, ws, widths: dict):
        for col, w in widths.items():
            ws.column_dimensions[col].width = w

    def title(self, ws, text: str, sub: str = ""):
        ws["A1"] = text
        ws["A1"].font = font(bold=True, size=14, color=C_HEAD)
        if sub:
            ws["A2"] = sub
            ws["A2"].font = font(italic=True, color="595959")

    def fmt_for(self, field: dict) -> str:
        t = field.get("type")
        if t == "date":
            return FMT["date"]
        if t == "pct":
            return FMT["%"]
        if t == "int":
            return FMT["int"]
        if t == "text":
            return FMT["text"]
        return FMT.get(field.get("unit", ""), FMT["number"] if field.get("unit") is None else FMT["number"])

    def tbl_formula(self, tbl: str, expr: str) -> str:
        """Translate {Col} placeholders into table this-row references."""
        out = const_subst(expr)
        return re.sub(r"\{([A-Za-z_][A-Za-z0-9_]*)\}", lambda m: f"{tbl}[[#This Row],[{m.group(1)}]]", out)

    def a1_formula(self, expr: str, cols: dict, row: int) -> str:
        """Translate {Col} placeholders into A1 references on one row (conditional formats cannot use table references)."""
        out = const_subst(expr)
        return re.sub(r"\{([A-Za-z_][A-Za-z0-9_]*)\}", lambda m: f"${cols[m.group(1)]}{row}", out)

    # ---- tables ------------------------------------------------------------------
    def prefill_rows(self, tdef: dict) -> int:
        p = tdef.get("prefill")
        if p == "dates":
            return len(self.days)
        if p == "months":
            return 12
        if p == "prior_months":
            return 24   # prior year and reporting year (months before go-live)
        return int(tdef.get("rows", 100))

    def write_table(self, ws, tdef: dict, anchor_col: int = 1, anchor_row: int = 1, data: list[list] | None = None, style="TableStyleMedium2"):
        """Row anchor_row = labels, anchor_row+1 = header (field names), then data rows. Returns table info."""
        tbl = tdef["table"]
        fields = tdef["fields"]
        nrows = self.prefill_rows(tdef)
        hdr_row = anchor_row + 1
        first = hdr_row + 1
        last = first + nrows - 1
        cols = {}
        for j, f in enumerate(fields):
            c = anchor_col + j
            L = get_column_letter(c)
            cols[f["name"]] = L
            lab = f.get("label", f["name"]) + (f" ({f['unit']})" if f.get("unit") else "")
            lc = ws.cell(anchor_row, c, lab)
            lc.font = font(italic=True, color="595959", size=9)
            lc.fill = fill(C_LABEL)
            lc.alignment = Alignment(wrap_text=True, vertical="bottom")
            hc = ws.cell(hdr_row, c, f["name"])
            hc.font = font(bold=True, color="FFFFFF")
            hc.fill = fill(C_HEAD_CALC if f.get("calc") else C_HEAD)
            hc.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            ws.column_dimensions[L].width = max(11, min(28, len(lab) * 0.9 + 2)) if f.get("type") != "text" else 32
        ws.row_dimensions[anchor_row].height = 30
        # prefill / data
        p = tdef.get("prefill")
        key = fields[0]["name"]
        for i in range(nrows):
            r = first + i
            if p == "dates":
                ws.cell(r, anchor_col, self.days[i])
            elif p == "months":
                ws.cell(r, anchor_col, date(self.year, i + 1, 1))
            elif p == "prior_months":
                ws.cell(r, anchor_col, date(self.year - 1 + i // 12, i % 12 + 1, 1))
        if data:
            by_key = {}
            if p in ("dates", "months", "prior_months"):
                for row in data:
                    by_key[row[0]] = row
                for i in range(nrows):
                    k = ws.cell(first + i, anchor_col).value
                    row = by_key.get(k)
                    if row:
                        for j, v in enumerate(row):
                            if j == 0 or v is None:
                                continue
                            ws.cell(first + i, anchor_col + j, v)
            else:
                for i, row in enumerate(data[:nrows]):
                    for j, v in enumerate(row):
                        if v is not None:
                            ws.cell(first + i, anchor_col + j, v)
        # formats, formulas, validation
        for j, f in enumerate(fields):
            c = anchor_col + j
            L = get_column_letter(c)
            nf = self.fmt_for(f)
            calc = f.get("calc")
            for r in range(first, last + 1):
                cell = ws.cell(r, c)
                cell.number_format = nf
                if calc:
                    cell.value = "=" + self.tbl_formula(tbl, calc)
                    cell.font = font(color="595959")
                    self.formula_count += 1
            if not calc:
                dv = self.dv_for(f, tdef, L, first, last)
                if dv is not None:
                    ws.add_data_validation(dv)
                    dv.add(f"{L}{first}:{L}{last}")
            # conditional highlights (schema-driven): a value required under a condition, a warning condition,
            # and a duplicated key in a prefilled table (paste bypasses data validation)
            rules = []
            if f.get("required_when"):
                rules.append(f'AND({self.a1_formula(f["required_when"], cols, first)},${L}{first}="")')
            if f.get("warn_when"):
                rules.append(self.a1_formula(f["warn_when"], cols, first))
            if f.get("key") and p in ("dates", "months", "prior_months"):
                rules.append(f"COUNTIF(${L}${first}:${L}${last},${L}{first})>1")
            for rule in rules:
                ws.conditional_formatting.add(f"{L}{first}:{L}{last}", FormulaRule(formula=[rule], fill=fill(C_WARN), font=Font(color="9C0006")))
        ref = f"{get_column_letter(anchor_col)}{hdr_row}:{get_column_letter(anchor_col + len(fields) - 1)}{last}"
        t = Table(displayName=tbl, ref=ref)
        t.tableStyleInfo = TableStyleInfo(name=style, showRowStripes=True, showFirstColumn=False, showLastColumn=False, showColumnStripes=False)
        t._initialise_columns()
        for col, f in zip(t.tableColumns, fields):
            col.name = f["name"]
            if f.get("calc"):
                col.calculatedColumnFormula = TableFormula(attr_text=self.tbl_formula(tbl, f["calc"]))
        ws.add_table(t)
        info = {"sheet": ws.title, "hdr_row": hdr_row, "first": first, "last": last, "cols": cols, "fields": fields}
        self.tables[tbl] = info
        return info

    def dv_for(self, f: dict, tdef: dict | None = None, L: str = "", first: int = 0, last: int = 0):
        t = f.get("type")
        msg = f.get("desc", "")
        prefill = (tdef or {}).get("prefill")
        if f.get("key") and prefill in ("dates", "months", "prior_months") and L:
            # prefilled key column: the only realistic error is a duplicated day or month, which SUMIFS would double count
            dv = DataValidation(type="custom", formula1=f"COUNTIF(${L}${first}:${L}${last},{L}{first})=1", allow_blank=True,
                                showErrorMessage=True, errorStyle="stop", showInputMessage=bool(msg))
            dv.errorTitle = "Duplicate date"
            dv.error = "This date already has a row in the table. Enter the values on the existing row."
        elif t == "list" and f.get("list"):
            dv = DataValidation(type="list", formula1=f"=lst_{f['list']}", allow_blank=True, showErrorMessage=True, showInputMessage=bool(msg))
            dv.errorTitle = "Not in list"
            dv.error = f"Choose a value from the list (edit the Lists sheet to add one)."
        elif t in ("number", "pct", "int"):
            lo, hi = f.get("min", 0), f.get("max", 1e12)
            dv = DataValidation(type="whole" if t == "int" else "decimal", operator="between", formula1=num_text(lo), formula2=num_text(hi),
                                allow_blank=True, showErrorMessage=True, errorStyle="warning", showInputMessage=bool(msg))
            dv.errorTitle = "Outside expected range"
            dv.error = f"Expected between {lo:,.15g} and {hi:,.15g}. Click Yes to keep the value if it is correct."
        elif t == "date":
            # rows dated outside the year have no Calc_Daily row and would vanish from every KPI
            lo = "=EDATE(cfg_YearStart,-12)" if prefill == "prior_months" else "=cfg_YearStart"
            dv = DataValidation(type="date", operator="between", formula1=lo, formula2="=cfg_YearEnd", allow_blank=True,
                                showErrorMessage=True, errorStyle="warning", showInputMessage=bool(msg))
            if prefill == "prior_months":
                dv.errorTitle = "Month outside the prior or reporting year"
                dv.error = "Months must fall in the prior year or the reporting year (the rolling 12-month window). Click No and correct the month."
            else:
                dv.errorTitle = "Date outside the reporting year"
                dv.error = "Dates must fall in the reporting year: rows outside it are not counted in any KPI. Click No and correct the date."
        else:
            return None
        if msg:
            dv.promptTitle = f.get("label", f["name"])[:32]
            dv.prompt = msg[:250]
        return dv

    # ---- sheets ------------------------------------------------------------------
    def build(self):
        site = self.s["site"]
        wb = self.wb
        # order: README, reports, inputs, config/lists, calc
        ws_readme = wb.create_sheet("README")
        ws_rep = wb.create_sheet("Daily_Report")
        ws_dash = wb.create_sheet("Dashboard")
        ws_month = wb.create_sheet("Monthly_Summary")
        input_sheets = {}
        for tdef in self.s["tables"]:
            if tdef["sheet"] in ("Config", "Lists"):
                continue
            if tdef["sheet"] not in input_sheets:
                input_sheets[tdef["sheet"]] = wb.create_sheet(tdef["sheet"])
        ws_cfg = wb.create_sheet("Config")
        ws_lists = wb.create_sheet("Lists")
        ws_cd = wb.create_sheet("Calc_Daily")
        ws_chart = wb.create_sheet("Chart_Data")
        ws_dict = wb.create_sheet("Data_Dictionary")

        sample = self.make_sample() if self.sample_days else {}
        self.sample = sample

        self.sheet_lists(ws_lists)
        self.sheet_config(ws_cfg, sample)
        for tdef in self.s["tables"]:
            if tdef["sheet"] in input_sheets:
                self.sheet_input(input_sheets[tdef["sheet"]], tdef, sample.get(tdef["table"]))
        self.sheet_calc_daily(ws_cd)
        self.sheet_chart_data(ws_chart)
        self.sheet_daily_report(ws_rep)
        self.sheet_monthly(ws_month)
        self.sheet_dashboard(ws_dash, ws_chart)
        self.sheet_dictionary(ws_dict)
        self.sheet_readme(ws_readme)

        for ws in (ws_rep, ws_dash, ws_month):
            ws.sheet_properties.tabColor = TAB["report"]
        for ws in input_sheets.values():
            ws.sheet_properties.tabColor = TAB["input"]
        for ws in (ws_cfg, ws_lists):
            ws.sheet_properties.tabColor = TAB["config"]
        for ws in (ws_cd, ws_chart, ws_dict):
            ws.sheet_properties.tabColor = TAB["calc"]
        wb.active = wb.sheetnames.index("Daily_Report")
        wb.calculation.fullCalcOnLoad = True
        wb.properties.title = f"{site['name']} Production System {self.year}"
        wb.properties.creator = "Mako Production System"
        wb.properties.subject = "Daily production reporting"
        self.check_catalogue()
        return wb

    # ---- consistency of the KPI catalogue and report formulas with the schema ----------
    def check_catalogue(self):
        """Every tblX[Col] reference, quoted list criterion and budget column in the KPI catalogue and in the generated
        formulas must exist in the schema; otherwise Excel shows #REF! or a renamed list value silently sums to zero."""
        s = self.s
        tables = {t["table"]: {f["name"]: f for f in t["fields"]} for t in s["tables"]}
        lists = dict(s["lists"])
        lists.update(s.get("derived_lists", {}))
        for rdef in s["reference_tables"].values():
            tables[rdef["table"]] = {c["name"]: c for c in rdef["columns"]}
            lists[rdef["key"]] = [row[0] for row in rdef["rows"]]
        for lname in lists:
            tables[f"tblList_{lname}"] = {lname: {"name": lname, "type": "text"}}
        ref_re = re.compile(r"(tbl\w+)\[(?:\[#This Row\],)?\[?(\w+)\]?\]?")
        crit_re = re.compile(r"(tbl\w+)\[(\w+)\],\"([^\"<>=?]+)\"")  # skips operators and the ? sentinel
        problems = []

        def check(text: str, where: str):
            for tbl, col in ref_re.findall(text):
                if tbl not in tables:
                    problems.append(f"{where}: table {tbl} is not in the schema")
                elif col not in tables[tbl]:
                    problems.append(f"{where}: column {tbl}[{col}] is not in the schema")
            for tbl, col, val in crit_re.findall(text):
                f = tables.get(tbl, {}).get(col)
                lname = (f or {}).get("list") or (f or {}).get("values_from")
                if lname and val not in lists.get(lname, []):
                    problems.append(f"{where}: value \"{val}\" is not in list {lname} ({tbl}[{col}])")

        budget_cols = tables["tblBudget"]
        for k in KPIS:
            for part in ("day", "num", "den"):
                if k.get(part):
                    check(k[part], f"KPI {k['key']}")
            if k.get("budget") and k["budget"] not in budget_cols:
                problems.append(f"KPI {k['key']}: budget column tblBudget[{k['budget']}] is not in the schema")
            for part in ("num", "den"):
                for key in re.findall(r"\{(\w+)\}", k.get(part, "")):
                    if key not in KPI:
                        problems.append(f"KPI {k['key']}: {part} refers to unknown KPI {key}")
        for ws in (self.wb["Daily_Report"], self.wb["Config"], self.wb["Monthly_Summary"], self.wb["Lists"]):
            for row in ws.iter_rows():
                for c in row:
                    if isinstance(c.value, str) and c.value.startswith("="):
                        check(c.value, f"{ws.title}!{c.coordinate}")
        for ws_name in ("Calc_Daily", "Chart_Data"):
            for c in self.wb[ws_name][2]:
                if isinstance(c.value, str) and c.value.startswith("="):
                    check(c.value, f"{ws_name}!{c.coordinate}")
        for tbl, info in self.tables.items():
            ws = self.wb[info["sheet"]]
            for L in info["cols"].values():
                v = ws[f"{L}{info['first']}"].value
                if isinstance(v, str) and v.startswith("="):
                    check(v, f"{tbl} calc column {L}")
        if problems:
            raise ValueError("schema / catalogue mismatch:\n  " + "\n  ".join(problems[:20]))

    # Lists ------------------------------------------------------------------------
    def sheet_lists(self, ws):
        self.title(ws, "Lists", "Drop-down values. Add rows at the bottom of a list; do not rename headers. Every list is referenced by name (lst_...) in data validation.")
        col = 1
        row0 = 4
        all_lists = dict(self.s["lists"])
        all_lists.update(self.s.get("derived_lists", {}))
        for lname, values in all_lists.items():
            tdef = {"table": f"tblList_{lname}", "fields": [{"name": lname, "label": lname.replace("_", " "), "type": "text"}], "rows": len(values)}
            info = self.write_table(ws, tdef, anchor_col=col, anchor_row=row0, data=[[v] for v in values], style="TableStyleLight9")
            L = get_column_letter(col)
            self.name(f"lst_{lname}", self.dyn_range("Lists", L, info["first"]))
            ws.column_dimensions[L].width = max(16, max(len(v) for v in values) + 2)
            col += 2
        col += 1
        for rname, rdef in self.s["reference_tables"].items():
            if rdef["sheet"] != "Lists":
                continue
            tdef = {"table": rdef["table"], "fields": rdef["columns"], "rows": len(rdef["rows"])}
            info = self.write_table(ws, tdef, anchor_col=col, anchor_row=row0, data=rdef["rows"], style="TableStyleLight9")
            keyL = info["cols"][rdef["key"]]
            self.name(f"lst_{rdef['key']}", self.dyn_range("Lists", keyL, info["first"]))
            ws.column_dimensions[keyL].width = 26
            col += len(rdef["columns"]) + 1
        ws.freeze_panes = "A6"

    # Config -----------------------------------------------------------------------
    def sheet_config(self, ws, sample):
        site = self.s["site"]
        self.title(ws, "Config", "Site settings and opening balances. Names in column C are used by the formulas.")
        consts = self.s["constants"]
        settings = [
            ("Site", site["name"], "cfg_Site", "text"),
            ("Company", site["company"], "cfg_Company", "text"),
            ("Report title", f"{site['name']} daily production report", "cfg_ReportTitle", "text"),
            ("Date of last LTI before this year", sample.get("prior_lti") if sample else None, "cfg_PriorLTIDate", "date"),
        ] + [(label, consts[key], nm, "number" if isinstance(consts[key], float) else "count") for key, (nm, label) in CONSTANTS.items()] + [
            # calculated: the reporting year is fixed when the workbook is generated (dates in Calc_Daily and the prefilled tables)
            ("Calculated (do not type here)", None, "", "head"),
            ("Reporting year (from the generator; run build_workbook.py --year to change)", "=YEAR(INDEX(cd_Date,1))", "cfg_Year", "int"),
            ("Year start", "=DATE(cfg_Year,1,1)", "cfg_YearStart", "date"),
            ("Year end", "=DATE(cfg_Year,12,31)", "cfg_YearEnd", "date"),
            ("Last date with plant data", '=SUMPRODUCT(MAX((tblPlant[Milled_t]<>"")*tblPlant[Date]))', "cfg_LastDataDate", "date"),
            ("Plant feed stockpile check", '=IF(COUNTIF(tblStockpiles[Plant_Feed],"Y")=1,"OK","Exactly one stockpile must have Plant_Feed = Y")', "cfg_FeedCheck", "text"),
            ("Year check", '=IF(cfg_Year<>YEAR(INDEX(cd_Date,1)),"Year mismatch: rebuild the workbook with build_workbook.py --year","")', "cfg_YearCheck", "text"),
        ]
        r = 4
        ws.cell(r, 1, "Setting").font = font(bold=True); ws.cell(r, 2, "Value").font = font(bold=True); ws.cell(r, 3, "Name").font = font(bold=True)
        for lab, val, nm, typ in settings:
            r += 1
            if typ == "head":
                r += 1
                ws.cell(r, 1, lab).font = font(bold=True, color="595959")
                continue
            ws.cell(r, 1, lab)
            c = ws.cell(r, 2, val)
            c.number_format = {"date": FMT["date"], "int": "0", "count": "#,##0", "number": "0.0000"}.get(typ, "@")
            c.font = font(color="1F3864" if not (isinstance(val, str) and val.startswith("=")) else "595959")
            c.border = BORDER
            ws.cell(r, 3, nm).font = font(color="7F7F7F", size=9)
            self.name(nm, f"Config!$B${r}")
            if nm in ("cfg_FeedCheck", "cfg_YearCheck"):
                ok = '"OK"' if nm == "cfg_FeedCheck" else '""'
                ws.conditional_formatting.add(f"B{r}", FormulaRule(formula=[f"B{r}<>{ok}"], fill=fill(C_WARN), font=Font(color="9C0006", bold=True)))
        self.set_widths(ws, {"A": 60, "B": 30, "C": 18})
        # reference tables on Config (stockpiles, with spare rows for stockpiles opened during the year) and safety history
        col = 5
        for rname, rdef in self.s["reference_tables"].items():
            if rdef["sheet"] != "Config":
                continue
            rows = sample.get(rdef["table"], rdef["rows"]) if sample else rdef["rows"]
            tdef = {"table": rdef["table"], "title": "Stockpile opening balances: surveyed tonnes at the start of Survey_Date (1 January if blank); spare rows for stockpiles opened later",
                    "fields": rdef["columns"], "rows": len(rows) + int(rdef.get("spare_rows", 0))}
            ws.cell(3, col, tdef["title"]).font = font(bold=True, color=C_HEAD)
            info = self.write_table(ws, tdef, anchor_col=col, anchor_row=4, data=rows, style="TableStyleLight9")
            col += len(rdef["columns"]) + 1
        for tdef in self.s["tables"]:
            if tdef["sheet"] == "Config":
                ws.cell(3, col, tdef["title"]).font = font(bold=True, color=C_HEAD)
                self.write_table(ws, tdef, anchor_col=col, anchor_row=4, data=sample.get(tdef["table"]) if sample else None, style="TableStyleLight9")
                col += len(tdef["fields"]) + 1

    # Input sheets -----------------------------------------------------------------
    def sheet_input(self, ws, tdef, data):
        if ws["A1"].value is None:
            self.title(ws, tdef["title"], f"Owner: {tdef.get('owner', '')}. Type only in the blue-headed columns; grey-headed columns are formulas. "
                                          "Add rows at the bottom only; never insert columns. Paste with Paste Special > Values only. Red cells need attention (missing or unknown value, duplicate date).")
            anchor_row = 4
        else:
            anchor_row = ws.max_row + 3
        info = self.write_table(ws, tdef, anchor_col=1, anchor_row=anchor_row, data=data)
        ws.freeze_panes = ws.cell(info["first"], 2)

    # Calc_Daily -------------------------------------------------------------------
    def plan_calc_columns(self):
        """Assign Calc_Daily columns. Returns ordered list of (header, kind, kpi, period, is_budget)."""
        cols = [("Date", "base"), ("Month", "base"), ("Day", "base"), ("Days_In_Month", "base"), ("Month_Start", "base")]
        for k in KPIS:
            kind = k["kind"]
            if kind in ("sum", "ratio"):
                cols += [(k["key"], "actual"), (k["key"] + "_MTD", "actual"), (k["key"] + "_YTD", "actual")]
            else:
                cols += [(k["key"], "actual")]
        for k in KPIS:
            if self.has_budget(k):
                cols += [(f"Bud_{k['key']}_Month", "budget"), (f"Bud_{k['key']}_MTD", "budget"), (f"Bud_{k['key']}_YTD", "budget")]
        self.cd_cols = {h: get_column_letter(i + 1) for i, (h, _) in enumerate(cols)}
        return cols

    def has_budget(self, k) -> bool:
        if k.get("budget"):
            return True
        if k["kind"] == "ratio":
            refs = {x for x in re.findall(r"\{([A-Za-z_][A-Za-z0-9_]*)\}", k["num"] + k["den"]) if x in KPI}
            return bool(refs) and all(self.has_budget(KPI[x]) for x in refs)
        return False

    def resolve(self, expr: str, r: int, period: str = "", budget: bool = False) -> str:
        c = self.cd_cols

        def rep(m):
            key = m.group(1)
            if key == "d":
                return f"$A{r}"
            if key == "ms":
                return f"$E{r}"
            if key == "dim":
                return f"$D{r}"
            if key == "r":
                return str(r)
            if key in PLACEHOLDERS:
                return PLACEHOLDERS[key]
            if key in KPI:
                if budget:
                    hdr = f"Bud_{key}_{period}"
                else:
                    hdr = key if period in ("", "Day") else f"{key}_{period}"
                return f"${c[hdr]}{r}"
            raise KeyError(f"unknown placeholder {{{key}}} in {expr}")

        return re.sub(r"\{([A-Za-z_][A-Za-z0-9_]*)\}", rep, expr)

    def sheet_calc_daily(self, ws):
        cols = self.plan_calc_columns()
        ws.append([h for h, _ in cols])
        for j, (h, kind) in enumerate(cols, start=1):
            c = ws.cell(1, j)
            c.font = font(bold=True, color="FFFFFF", size=9)
            c.fill = fill(C_HEAD if kind != "budget" else C_HEAD_CALC)
            c.alignment = Alignment(wrap_text=True, vertical="center")
            ws.column_dimensions[get_column_letter(j)].width = 12
        ws.row_dimensions[1].height = 42
        n = len(self.days)
        first, last = 2, n + 1
        c = self.cd_cols
        for i, d in enumerate(self.days):
            r = first + i
            ws.cell(r, 1, d).number_format = FMT["date"]
            ws.cell(r, 2, f"=MONTH($A{r})")
            ws.cell(r, 3, f"=DAY($A{r})")
            ws.cell(r, 4, f"=DAY(EOMONTH($A{r},0))")
            ws.cell(r, 5, f"=DATE(YEAR($A{r}),MONTH($A{r}),1)").number_format = FMT["date"]
            for k in KPIS:
                key, kind = k["key"], k["kind"]
                nf = FMT.get(k["unit"], FMT["number"])
                if kind == "sum":
                    day = "=" + self.resolve(k["day"], r)
                    mtd = f"=${c[key]}{r}" if r == first else f"=IF($B{r}=$B{r-1},${c[key + '_MTD']}{r-1},0)+${c[key]}{r}"
                    ytd = f"=${c[key]}{r}" if r == first else f"=${c[key + '_YTD']}{r-1}+${c[key]}{r}"
                    for hdr, f_ in ((key, day), (key + "_MTD", mtd), (key + "_YTD", ytd)):
                        cell = ws[f"{c[hdr]}{r}"]; cell.value = f_; cell.number_format = nf
                elif kind == "ratio":
                    for period in ("Day", "MTD", "YTD"):
                        hdr = key if period == "Day" else f"{key}_{period}"
                        num = self.resolve(k["num"], r, period)
                        den = self.resolve(k["den"], r, period)
                        cell = ws[f"{c[hdr]}{r}"]
                        cell.value = f'=IFERROR(IF(({den})>0,({num})/({den}),""),"")'
                        cell.number_format = nf
                else:
                    cell = ws[f"{c[key]}{r}"]
                    cell.value = "=" + self.resolve(k["day"], r)
                    cell.number_format = nf
            for k in KPIS:
                if not self.has_budget(k):
                    continue
                key = k["key"]
                nf = FMT.get(k["unit"], FMT["number"])
                if k.get("budget"):
                    bcol = k["budget"]
                    month = f"=SUMIFS(tblBudget[{bcol}],tblBudget[Month],$E{r})"
                    mtd = f"=${c[f'Bud_{key}_Month']}{r}*$C{r}/$D{r}"
                    ytd = f'=SUMIFS(tblBudget[{bcol}],tblBudget[Month],"<"&$E{r})+${c[f"Bud_{key}_MTD"]}{r}'
                    for period, f_ in (("Month", month), ("MTD", mtd), ("YTD", ytd)):
                        cell = ws[f"{c[f'Bud_{key}_{period}']}{r}"]; cell.value = f_; cell.number_format = nf
                else:  # ratio budget
                    for period in ("Month", "MTD", "YTD"):
                        num = self.resolve(k["num"], r, period, budget=True)
                        den = self.resolve(k["den"], r, period, budget=True)
                        cell = ws[f"{c[f'Bud_{key}_{period}']}{r}"]
                        cell.value = f'=IFERROR(IF(({den})>0,({num})/({den}),""),"")'
                        cell.number_format = nf
        self.formula_count += (last - first + 1) * (len(cols) - 1)
        for h, _ in cols:
            L = c[h]
            self.name(f"cd_{h}", f"Calc_Daily!${L}${first}:${L}${last}")
        ws.freeze_panes = "B2"
        ws.protection.sheet = True
        self.cd_rows = (first, last)

    # Chart_Data -------------------------------------------------------------------
    def sheet_chart_data(self, ws):
        first, last = self.cd_rows
        series = [
            ("Date", None, False),
            ("Milled_t", "Milled_t", True),
            ("Budget_Milled_daily", "Bud_Milled_t_Month/DIM", False),
            ("Gold_Poured_YTD", "Gold_Poured_oz_YTD", True),
            ("Budget_Gold_Poured_YTD", "Bud_Gold_Poured_oz_YTD", False),
            ("Ore_Mined_t", "Ore_Mined_t", True),
            ("Waste_t", "Waste_t", True),
            ("Recovery_MTD", "Recovery_pct_MTD", True),
            ("Head_Grade_MTD", "Head_Grade_gpt_MTD", True),
            ("TRIFR_12m", "TRIFR_12m", True),
            ("Budget_Recovery_MTD", "Bud_Recovery_pct_MTD", False),
        ]
        ws.append([s[0] for s in series])
        for j in range(1, len(series) + 1):
            ws.cell(1, j).font = font(bold=True, color="FFFFFF", size=9); ws.cell(1, j).fill = fill(C_HEAD)
            ws.column_dimensions[get_column_letter(j)].width = 14
        c = self.cd_cols
        for r in range(first, last + 1):
            ws.cell(r, 1, f"=Calc_Daily!$A{r}").number_format = FMT["date"]
            for j, (hdr, src, guard) in enumerate(series[1:], start=2):
                if src == "Bud_Milled_t_Month/DIM":
                    val = f"Calc_Daily!${c['Bud_Milled_t_Month']}{r}/Calc_Daily!$D{r}"
                else:
                    val = f"Calc_Daily!${c[src]}{r}"
                # beyond the last data date, or where the KPI is blank text (no assay yet), plot a gap rather than zero
                f_ = f"=IF(OR($A{r}>cfg_LastDataDate,NOT(ISNUMBER({val}))),NA(),{val})" if guard else f"={val}"
                ws.cell(r, j, f_)
        self.formula_count += (last - first + 1) * len(series)
        ws.protection.sheet = True
        self.chart_series = series

    # Daily_Report -----------------------------------------------------------------
    def sheet_daily_report(self, ws):
        c = self.cd_cols
        ws.sheet_view.showGridLines = False
        self.set_widths(ws, {"A": 2, "B": 34, "C": 14, "D": 13, "E": 13, "F": 13, "G": 9, "H": 13, "I": 13, "J": 9, "K": 2, "L": 10})
        ws["B2"] = "=cfg_ReportTitle"; ws["B2"].font = font(bold=True, size=16, color=C_HEAD)
        ws["B3"] = "=cfg_Company&\" | \"&cfg_Site"; ws["B3"].font = font(italic=True, color="595959")
        ws["B5"] = "Report date"; ws["B5"].font = font(bold=True)
        ws["C5"] = self.report_date if self.report_date else DEFAULT_DATE_FORMULA
        ws["C5"].number_format = "ddd dd-mmm-yyyy"; ws["C5"].font = font(bold=True, color=C_HEAD); ws["C5"].fill = fill(C_KEY); ws["C5"].border = BORDER
        ws["C5"].protection = Protection(locked=False)
        ws.merge_cells("C5:E5")
        # the instruction lives in the input message (not printed); a date outside the year is stopped, and F5 warns if one slips through
        dv = DataValidation(type="date", operator="between", formula1="=cfg_YearStart", formula2="=cfg_YearEnd", allow_blank=False,
                            showErrorMessage=True, errorStyle="stop", showInputMessage=True)
        dv.errorTitle = "Date outside the reporting year"
        dv.error = "The report date must fall in the reporting year (1 January to 31 December)."
        dv.promptTitle = "Report date"
        dv.prompt = ("Type a date in the reporting year to change the report day. To restore the default (yesterday, or the last day "
                     f"with plant data) enter {DEFAULT_DATE_FORMULA}")[:255]
        ws.add_data_validation(dv); dv.add("C5")
        ws["F5"] = '=IF(ISNUMBER(rpt_Row),cfg_YearCheck,"Report date is outside the reporting year: enter a date between 1 January and 31 December "&cfg_Year)'
        ws["F5"].font = font(bold=True, size=9, color="C00000")
        ws["L5"] = "=IFERROR(MATCH($C$5,cd_Date,0),NA())"; ws["L5"].font = font(color="BFBFBF", size=8)
        ws["L4"] = "row"; ws["L4"].font = font(color="BFBFBF", size=8)
        self.name("rpt_Date", "Daily_Report!$C$5")
        self.name("rpt_Row", "Daily_Report!$L$5")
        ws["B6"] = ('=IF(ISNUMBER(rpt_Row),"Data to "&IF(N(cfg_LastDataDate)>0,' + xl_date_text("cfg_LastDataDate") + ',"(no plant data yet)")'
                    '&"  |  Week "&_xlfn.ISOWEEKNUM($C$5)&"  |  Day "&DAY($C$5)&" of "&DAY(EOMONTH($C$5,0)),"Report date is outside the reporting year")')
        ws["B6"].font = font(size=9, color="595959")
        # completeness and integrity line: rows entered for the report date per table, and movement rows no KPI can count
        ws["B7"] = ('="Rows for this date: Plant "&COUNTIFS(tblPlant[Date],rpt_Date,tblPlant[Milled_t],"<>")'
                    '&", Movements "&COUNTIFS(tblMovement[Date],rpt_Date)&", Fleet "&COUNTIFS(tblFleet[Date],rpt_Date)'
                    '&", Mining daily "&COUNTIFS(tblMiningDaily[Date],rpt_Date,tblMiningDaily[Diesel_L],"<>")'
                    '&", Safety "&COUNTIFS(tblSafety[Date],rpt_Date,tblSafety[Hours_Employees],"<>")&", Comments "&COUNTIFS(tblCommentary[Date],rpt_Date)'
                    '&"  |  Movement rows with unknown source or destination: "&(COUNTIF(tblMovement[Source_Type],"?")+COUNTIF(tblMovement[Dest_Type],"?"))'
                    '&"  |  Movement rows dated outside the year: "&(COUNTIFS(tblMovement[Date],"<"&cfg_YearStart)+COUNTIFS(tblMovement[Date],">"&cfg_YearEnd))')
        ws["B7"].font = font(size=9, color="595959")
        self.name("rpt_Checks", "Daily_Report!$B$7")

        headers = ["", "Unit", "Day", "MTD", "MTD budget", "Var %", "YTD", "YTD budget", "Var %"]
        r = 8
        break_row = None
        for section, items in REPORT_LAYOUT:
            if section == PAGE_BREAK_BEFORE:
                break_row = r - 1
            for j, h in enumerate(headers):
                cell = ws.cell(r, 2 + j, section if j == 0 else h)
                cell.font = font(bold=True, color="FFFFFF"); cell.fill = fill(C_HEAD)
                cell.alignment = Alignment(horizontal="left" if j == 0 else "center")
            r += 1
            for key, show_budget in items:
                k = KPI[key]
                nf = FMT.get(k["unit"], FMT["number"])
                ws.cell(r, 2, k["label"]).font = font()
                # rates print their hours basis (per 1,000,000 h) so the reader knows the convention
                unit = '="per "&TEXT(cfg_RateBasisHours,"#,##0")&" h"' if k["unit"] == "rate" else ("" if k["unit"] in ("int", "ratio", "days") else k["unit"])
                ws.cell(r, 3, unit).font = font(color="7F7F7F", size=9)
                has_periods = k["kind"] in ("sum", "ratio")
                def idx(hdr):
                    return f'=IFERROR(INDEX(cd_{hdr},rpt_Row),"")'
                ws.cell(r, 4, idx(key)).number_format = nf
                if has_periods:
                    ws.cell(r, 5, idx(key + "_MTD")).number_format = nf
                    ws.cell(r, 8, idx(key + "_YTD")).number_format = nf
                if show_budget and self.has_budget(k):
                    ws.cell(r, 6, idx(f"Bud_{key}_MTD")).number_format = nf
                    ws.cell(r, 9, idx(f"Bud_{key}_YTD")).number_format = nf
                    ws.cell(r, 7, f'=IF(AND(ISNUMBER(F{r}),ISNUMBER(E{r}),F{r}<>0),E{r}/F{r}-1,"")').number_format = "+0.0%;-0.0%;0.0%"
                    ws.cell(r, 10, f'=IF(AND(ISNUMBER(I{r}),ISNUMBER(H{r}),I{r}<>0),H{r}/I{r}-1,"")').number_format = "+0.0%;-0.0%;0.0%"
                for j in range(2, 11):
                    ws.cell(r, j).border = BORDER
                    if j >= 4:
                        ws.cell(r, j).alignment = Alignment(horizontal="right")
                r += 1
            r += 1
        # stockpiles: one report row per tblStockpiles row including the spare rows, blank until a stockpile is named
        sp = self.tables["tblStockpiles"]
        n_sp = sp["last"] - sp["first"] + 1
        ws.cell(r - 1, 2, '=IF(cfg_FeedCheck="OK","",cfg_FeedCheck&" (Config): the plant feed deduction below is wrong until this is fixed")').font = font(bold=True, size=9, color="C00000")
        hdr = ["STOCKPILES (to report date)", "", "Opening t", "Opening g/t", "In t", "Out t", "Plant feed t", "Balance t", "Balance g/t"]
        for j, h in enumerate(hdr):
            cell = ws.cell(r, 2 + j, h); cell.font = font(bold=True, color="FFFFFF"); cell.fill = fill(C_HEAD)
            cell.alignment = Alignment(horizontal="left" if j == 0 else "center")
        r += 1
        # balance window: from the later of 1 January and the stockpile's Survey_Date (the opening balance is the surveyed
        # tonnage at the start of that day) to the report date; rows before it are not part of the balance
        for i in range(1, n_sp + 1):
            nm = f"INDEX(tblStockpiles[Stockpile],{i})"
            since = f'">="&MAX(cfg_YearStart,N(INDEX(tblStockpiles[Survey_Date],{i})))'
            in_win = f'tblMovement[Date],{since},tblMovement[Date],"<="&rpt_Date'
            pl_win = f'tblPlant[Date],{since},tblPlant[Date],"<="&rpt_Date'
            g = f'=IF(OR(NOT(ISNUMBER(rpt_Row)),{nm}=""),"",'   # blank when no stockpile is named or the report date is outside the year
            ws.cell(r, 2, f"{g}{nm})")
            ws.cell(r, 4, f"{g}INDEX(tblStockpiles[Opening_t],{i}))").number_format = FMT["t"]
            ws.cell(r, 5, f"{g}INDEX(tblStockpiles[Opening_gpt],{i}))").number_format = FMT["g/t"]
            ws.cell(r, 6, f'{g}SUMIFS(tblMovement[Tonnes_t],tblMovement[Destination],{nm},{in_win}))').number_format = FMT["t"]
            ws.cell(r, 7, f'{g}SUMIFS(tblMovement[Tonnes_t],tblMovement[Source],{nm},{in_win}))').number_format = FMT["t"]
            ws.cell(r, 8, f'{g}IF(INDEX(tblStockpiles[Plant_Feed],{i})="Y",SUMIFS(tblPlant[Milled_t],{pl_win})-SUMIFS(tblMovement[Tonnes_t],tblMovement[Dest_Type],"Plant",{in_win}),0))').number_format = FMT["t"]
            ws.cell(r, 9, f'=IF($B{r}="","",D{r}+F{r}-G{r}-H{r})').number_format = FMT["t"]
            # ounces: opening + in - out - plant feed (feed oz less direct tip oz)
            oz = (f'(D{r}*E{r}/{TROY}'
                  f'+SUMIFS(tblMovement[Contained_oz],tblMovement[Destination],{nm},{in_win})'
                  f'-SUMIFS(tblMovement[Contained_oz],tblMovement[Source],{nm},{in_win})'
                  f'-IF(INDEX(tblStockpiles[Plant_Feed],{i})="Y",SUMIFS(tblPlant[Feed_oz],{pl_win})-SUMIFS(tblMovement[Contained_oz],tblMovement[Dest_Type],"Plant",{in_win}),0))')
            ws.cell(r, 10, f'=IFERROR(IF(AND($B{r}<>"",I{r}>0),{oz}*{TROY}/I{r},""),"")').number_format = FMT["g/t"]
            for j in range(2, 11):
                ws.cell(r, j).border = BORDER
            r += 1
        r += 1
        # commentary
        cell = ws.cell(r, 2, "COMMENTARY"); cell.font = font(bold=True, color="FFFFFF"); cell.fill = fill(C_HEAD)
        for j in range(3, 11):
            ws.cell(r, j).fill = fill(C_HEAD)
        ws.cell(r, 3, "Area").font = font(bold=True, color="FFFFFF")
        r += 1
        for n in range(1, COMMENT_LINES + 1):
            ws.cell(r, 2, f'=IFERROR(INDEX(tblCommentary[Area],MATCH(rpt_Date&"|"&{n},tblCommentary[Key_Day],0)),"")').font = font(bold=True, size=9)
            ws.cell(r, 3, f'=IFERROR(INDEX(tblCommentary[Comment],MATCH(rpt_Date&"|"&{n},tblCommentary[Key_Day],0)),"")').font = font(size=9)
            ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=10)
            ws.cell(r, 3).alignment = Alignment(wrap_text=True, vertical="top")
            ws.row_dimensions[r].height = COMMENT_ROW_PT
            r += 1
        r += 1
        ws.cell(r, 2, '="Generated by the Mako Production System workbook. Printed "&' + xl_date_text("NOW()")
                + '&" "&TEXT(HOUR(NOW()),"00")&":"&TEXT(MINUTE(NOW()),"00")').font = font(italic=True, size=8, color="7F7F7F")
        # print setup
        # two portrait A4 pages: safety and mining, then processing, gold, stockpiles and commentary
        ws.print_area = f"B2:J{r}"
        ws.page_setup.orientation = "portrait"
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 2
        ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
        if break_row:
            ws.row_breaks.append(Break(id=break_row))
        ws.print_options.horizontalCentered = True
        ws.page_margins.left = ws.page_margins.right = 0.4
        ws.protection.sheet = True
        ws.protection.formatRows = False   # users may enlarge a commentary row for a long comment
        self.report_last_row = r

    # Monthly_Summary --------------------------------------------------------------
    def sheet_monthly(self, ws):
        ws.sheet_view.showGridLines = False
        self.set_widths(ws, {"A": 2, "B": 34, "C": 10, **{get_column_letter(i): 11 for i in range(4, 17)}})
        ws["B2"] = '=cfg_Site&" monthly summary "&cfg_Year'; ws["B2"].font = font(bold=True, size=16, color=C_HEAD)
        ws["B3"] = ('="Actuals to "&IF(N(cfg_LastDataDate)>0,' + xl_date_text("cfg_LastDataDate") + ',"(no plant data yet)")'
                    '&". Budget from the Budget sheet, pro-rata to that date for the current month and the Year column; full month otherwise. Ratios are month values, not averages of days."')
        ws["B3"].font = font(italic=True, size=9, color="595959")
        r = 5
        ws.cell(r, 2, "KPI").font = font(bold=True, color="FFFFFF"); ws.cell(r, 2).fill = fill(C_HEAD)
        ws.cell(r, 3, "").fill = fill(C_HEAD)
        for m in range(1, 13):
            cell = ws.cell(r, 3 + m, date(self.year, m, 1)); cell.number_format = "mmm"; cell.font = font(bold=True, color="FFFFFF"); cell.fill = fill(C_HEAD)
            cell.alignment = Alignment(horizontal="center")
        cell = ws.cell(r, 16, "Year"); cell.font = font(bold=True, color="FFFFFF"); cell.fill = fill(C_HEAD); cell.alignment = Alignment(horizontal="center")
        r += 1
        for key in MONTHLY_KPIS:
            k = KPI[key]
            nf = FMT.get(k["unit"], FMT["number"])
            has_b = self.has_budget(k)
            rows = [("Actual", "actual")] + ([("Budget", "budget"), ("Variance", "var")] if has_b else [])
            first_r = r
            for lab, kind in rows:
                ws.cell(r, 2, k["label"] if kind == "actual" else "").font = font(bold=(kind == "actual"))
                ws.cell(r, 3, lab).font = font(size=9, color="7F7F7F")
                for m in range(1, 13):
                    col = 3 + m
                    ms = f"DATE(cfg_Year,{m},1)"
                    me = f"MIN(EOMONTH({ms},0),cfg_LastDataDate)"
                    if kind == "actual":
                        src = f"cd_{key}_MTD" if k["kind"] in ("sum", "ratio") else f"cd_{key}"
                        f_ = f'=IF({ms}>cfg_LastDataDate,"",IFERROR(INDEX({src},MATCH({me},cd_Date,0)),""))'
                    elif kind == "budget":
                        # the month in progress compares MTD actual with the pro-rata budget to the same date
                        f_ = (f'=IF({ms}>cfg_LastDataDate,IFERROR(INDEX(cd_Bud_{key}_Month,MATCH({ms},cd_Date,0)),""),'
                              f'IFERROR(INDEX(cd_Bud_{key}_MTD,MATCH({me},cd_Date,0)),""))')
                    else:
                        f_ = f'=IF(AND(ISNUMBER({get_column_letter(col)}{first_r}),ISNUMBER({get_column_letter(col)}{first_r+1}),{get_column_letter(col)}{first_r+1}<>0),{get_column_letter(col)}{first_r}/{get_column_letter(col)}{first_r+1}-1,"")'
                    cell = ws.cell(r, col, f_)
                    cell.number_format = "+0.0%;-0.0%;0.0%" if kind == "var" else nf
                    cell.border = BORDER
                    if kind == "var":
                        cell.font = font(size=9, color="595959")
                # year column
                if kind == "actual":
                    src = f"cd_{key}_YTD" if k["kind"] in ("sum", "ratio") else f"cd_{key}"
                    f_ = f'=IFERROR(INDEX({src},MATCH(MIN(cfg_YearEnd,cfg_LastDataDate),cd_Date,0)),"")'
                elif kind == "budget":
                    # year to date budget at the last data date (full year while there is no data yet)
                    f_ = f'=IFERROR(INDEX(cd_Bud_{key}_YTD,MATCH(IF(N(cfg_LastDataDate)>0,MIN(cfg_YearEnd,cfg_LastDataDate),cfg_YearEnd),cd_Date,0)),"")'
                else:
                    f_ = f'=IF(AND(ISNUMBER(P{first_r}),ISNUMBER(P{first_r+1}),P{first_r+1}<>0),P{first_r}/P{first_r+1}-1,"")'
                cell = ws.cell(r, 16, f_); cell.number_format = "+0.0%;-0.0%;0.0%" if kind == "var" else nf; cell.border = BORDER; cell.font = font(bold=True)
                r += 1
            r += 1 if has_b else 0
        ws.freeze_panes = "D6"
        ws.page_setup.orientation = "landscape"; ws.page_setup.paperSize = ws.PAPERSIZE_A4; ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
        ws.protection.sheet = True

    # Dashboard --------------------------------------------------------------------
    def sheet_dashboard(self, ws, ws_chart):
        ws.sheet_view.showGridLines = False
        ws["B2"] = '=cfg_Site&" dashboard "&cfg_Year'; ws["B2"].font = font(bold=True, size=16, color=C_HEAD)
        ws["B3"] = '="Data to "&IF(N(cfg_LastDataDate)>0,' + xl_date_text("cfg_LastDataDate") + ',"(no plant data yet)")'; ws["B3"].font = font(italic=True, size=9, color="595959")
        first, last = self.cd_rows
        hdr = {s[0]: i + 1 for i, s in enumerate(self.chart_series)}

        def ref(col):
            return Reference(ws_chart, min_col=hdr[col], min_row=1, max_row=last)

        dates = Reference(ws_chart, min_col=1, min_row=first, max_row=last)

        def style(ch, title, y_title, height=8.5, width=24):
            ch.title = title; ch.height = height; ch.width = width
            ch.y_axis.title = y_title; ch.legend.position = "b"
            ch.x_axis.number_format = "mmm"; ch.x_axis.majorTimeUnit = "months"
            ch.y_axis.majorGridlines = None
            # openpyxl omits <c:delete/>; Excel then treats the axes as deleted and draws no labels or titles
            ch.x_axis.delete = False; ch.y_axis.delete = False
            ch.y_axis.crossAx = ch.x_axis.axId

        # charts are 24 cm wide: right-hand charts are anchored 15 default columns (25.4 cm) to the right of the left ones
        # 1 gold poured YTD vs budget
        ch1 = LineChart(); ch1.add_data(ref("Gold_Poured_YTD"), titles_from_data=True); ch1.add_data(ref("Budget_Gold_Poured_YTD"), titles_from_data=True)
        ch1.set_categories(dates); ch1.x_axis = DateAxis(crossAx=100); style(ch1, "Gold poured, year to date (oz)", "oz")
        ch1.series[1].graphicalProperties.line.dashStyle = "dash"
        ws.add_chart(ch1, "B5")
        # 2 milled daily vs budget daily
        # bar charts get a date axis too: a text axis over 365 daily categories repeats month labels
        ch2 = BarChart(); ch2.type = "col"; ch2.add_data(ref("Milled_t"), titles_from_data=True); ch2.set_categories(dates)
        ch2.x_axis = DateAxis(crossAx=100)
        ln = LineChart(); ln.add_data(ref("Budget_Milled_daily"), titles_from_data=True); ln.set_categories(dates)
        ln.x_axis = DateAxis(crossAx=100)   # same axId as the bar chart's axis, so one date axis is written
        ch2 += ln; style(ch2, "Milled tonnes per day vs budget", "t"); ch2.gapWidth = 30
        ws.add_chart(ch2, "Q5")
        # 3 ore and waste
        ch3 = BarChart(); ch3.type = "col"; ch3.grouping = "stacked"; ch3.overlap = 100
        ch3.add_data(ref("Ore_Mined_t"), titles_from_data=True); ch3.add_data(ref("Waste_t"), titles_from_data=True); ch3.set_categories(dates)
        ch3.x_axis = DateAxis(crossAx=100); style(ch3, "Ore and waste mined per day (t)", "t"); ch3.gapWidth = 30
        ws.add_chart(ch3, "B24")
        # 4 recovery MTD and head grade MTD
        ch4 = LineChart(); ch4.add_data(ref("Recovery_MTD"), titles_from_data=True); ch4.add_data(ref("Budget_Recovery_MTD"), titles_from_data=True); ch4.set_categories(dates)
        ch4.x_axis = DateAxis(crossAx=100); style(ch4, "Recovery, month to date (%)", "%"); ch4.y_axis.number_format = "0%"
        ch4.series[1].graphicalProperties.line.dashStyle = "dash"
        ch5 = LineChart(); ch5.add_data(ref("Head_Grade_MTD"), titles_from_data=True); ch5.set_categories(dates)
        ch5.y_axis.axId = 200; ch5.y_axis.title = "g/t"; ch5.y_axis.crosses = "max"; ch5.y_axis.majorGridlines = None
        ch5.y_axis.delete = False; ch5.x_axis.delete = True; ch5.x_axis.crossAx = 200   # secondary series share the date axis
        ch4 += ch5
        ws.add_chart(ch4, "Q24")
        # 5 TRIFR
        ch6 = LineChart(); ch6.add_data(ref("TRIFR_12m"), titles_from_data=True); ch6.set_categories(dates); ch6.x_axis = DateAxis(crossAx=100)
        style(ch6, "TRIFR, rolling 12 months", "per million hours")
        ws.add_chart(ch6, "B43")
        # print: one landscape A4 page wide (charts end around column AF, row 60)
        ws.print_area = "B2:AF60"
        ws.page_setup.orientation = "landscape"; ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
        ws.protection.sheet = True

    # Data_Dictionary --------------------------------------------------------------
    def sheet_dictionary(self, ws):
        self.title(ws, "Data dictionary", "Generated from mps_schema.json. The same definitions drive the SQL schema of the web application.")
        hdr = ["Table", "Sheet", "Field", "Label", "Type", "Unit", "Required", "Validation", "Description", "Formula"]
        r = 4
        for j, h in enumerate(hdr, start=1):
            c = ws.cell(r, j, h); c.font = font(bold=True, color="FFFFFF"); c.fill = fill(C_HEAD)
        for tdef in self.s["tables"]:
            for f in tdef["fields"]:
                r += 1
                val = f"list: {f['list']}" if f.get("list") else (f"{f.get('min', '')} to {f.get('max', '')}" if "min" in f or "max" in f else "")
                for j, v in enumerate([tdef["table"], tdef["sheet"], f["name"], f.get("label", ""), "calculated" if f.get("calc") else f["type"],
                                       f.get("unit", ""), "yes" if f.get("required") or f.get("key") else (f"when {f['required_when']}" if f.get("required_when") else ""),
                                       val, f.get("desc", ""), f.get("calc", "")], start=1):
                    ws.cell(r, j, v).font = font(size=9)
        r += 2
        ws.cell(r, 1, "KPI catalogue (Calc_Daily)").font = font(bold=True, color=C_HEAD)
        r += 1
        for j, h in enumerate(["Key", "Label", "Unit", "Kind", "Budget column", "Day formula / definition"], start=1):
            c = ws.cell(r, j, h); c.font = font(bold=True, color="FFFFFF"); c.fill = fill(C_HEAD)
        for k in KPIS:
            r += 1
            definition = k.get("day") or f"({k.get('num')}) / ({k.get('den')})"
            for j, v in enumerate([k["key"], k["label"], k["unit"], k["kind"], k.get("budget", "derived" if self.has_budget(k) else ""), definition], start=1):
                ws.cell(r, j, v).font = font(size=9)
        self.set_widths(ws, {"A": 18, "B": 16, "C": 24, "D": 30, "E": 11, "F": 7, "G": 8, "H": 18, "I": 60, "J": 60})
        ws.freeze_panes = "A5"
        ws.protection.sheet = True

    # README -----------------------------------------------------------------------
    def sheet_readme(self, ws):
        site = self.s["site"]
        tabs = ", ".join(dict.fromkeys(t["sheet"] for t in self.s["tables"] if t["sheet"] not in ("Config", "Lists")))
        ws.sheet_view.showGridLines = False
        ws.column_dimensions["A"].width = 3; ws.column_dimensions["B"].width = 120
        lines = [
            (f"{site['name']} Production System, {self.year}", "h1"),
            (f"{site['company']}. Version {self.s['version']}. Generated workbook: do not restructure by hand; change the schema and regenerate.", "sub"),
            ("", ""),
            ("How the workbook works", "h2"),
            (f"1. Inputs are the green tabs: {tabs}. Each is an Excel table. Type in the blue-headed columns only; grey-headed columns are formulas that fill automatically.", ""),
            ("2. Calc_Daily rebuilds every KPI per day from the tables (day, month to date, year to date, budget). Daily_Report, Monthly_Summary and Dashboard read only from Calc_Daily.", ""),
            ("3. There are no links to other workbooks and no macros. Copying the file anywhere keeps it working.", ""),
            ("4. Lists holds every drop-down. Add a new pit, stockpile or equipment class there; do not type free text into list columns.", ""),
            ("5. Config holds the site constants (troy ounce, hours per day, injury rate basis), the stockpile opening balances (exactly one stockpile has Plant_Feed = Y) and the monthly safety history used for rolling 12-month rates. The reporting year is fixed when the workbook is generated; to roll to a new year run build_workbook.py --year, do not type over it.", ""),
            ("6. If the report opens blank behind a yellow Protected View bar (a file received by e-mail or download), click Enable Editing; the workbook then calculates.", ""),
            ("", ""),
            ("Daily routine", "h2"),
            ("Plant: enter the row for the day (dates are pre-filled). Mining: add one movement row per source, material and destination; add the Mining_Daily row and one Fleet row per equipment class. HSE: enter the Safety row. Everyone: one Commentary row per topic. Then open Daily_Report, check the date in C5 and the 'Rows for this date' line in B7, and print or export to PDF.", ""),
            ("One person edits the file at a time: open, enter, save, close. If Excel opens it Read-Only someone else has it: wait, do not Save As a copy. Suggested order: Mining 06:30, Plant 06:45, HSE 07:00, Commentary 07:15. Simultaneous entry needs the file on SharePoint or OneDrive.", ""),
            ("", ""),
            ("Rules", "h2"),
            ("Do not insert columns inside tables, rename headers or move sheets. Add rows at the bottom of a table only. Paste only with Paste Special > Values; an ordinary paste removes the drop-downs and range checks. Blank means not reported; zero means reported as zero. Grades in g/t, tonnes dry metric, gold in troy ounces (31.1035 g).", ""),
            ("Ore versus waste follows the Material list: anything starting with Ore counts as ore. Ex-pit movements (Source type Pit) count to TMM and strip ratio; movements from stockpiles are rehandle. Every ore row needs a grade: ore tonnes without a grade are highlighted, excluded from the grade mined and shown as 'Ore mined without a grade' on the report. Enter either one daily total per source, material and destination or Day and Night rows, never both.", ""),
            ("Recovery is calculated from head and tails grades (1 minus tails over head). Gold recovered is milled tonnes x (head minus tails) / 31.1035. A day whose head or tails grade is not yet entered shows blank recovery and is left out of the period grades, recovery and gold recovered until the assay arrives; throughput uses only days with run hours. Monthly metallurgical accounting adjustments are handled outside this workbook until the web system takes over.", ""),
            ("Mill availability and utilisation count a day once any mill hours (run, planned or unplanned) are entered; a full-day shutdown is Planned maintenance 24 with Mill run hours 0.", ""),
            ("Stockpile balances on the report run from the later of 1 January and the stockpile's Survey_Date (Opening_t is the surveyed tonnage at the start of that day) to the report date. After a re-survey enter the new tonnes, grade and Survey_Date on Config. The plant feed stockpile is debited with milled tonnes less direct tip, so its balance includes ore tipped to the crusher but not yet milled.", ""),
            ("TRIFR and LTIFR are injuries per the rate basis on Config (1,000,000 hours) over the last 12 calendar months plus the current month to date. Months without daily Safety rows (the prior year, and this year before go-live) take their hours and injuries from the monthly history on Config.", ""),
            ("", ""),
            ("Support", "h2"),
            ("Schema and generator: mps_schema.json and build_workbook.py in the MPS tools folder. Report issues to the Director, Operations and Projects.", ""),
        ]
        r = 2
        for text, kind in lines:
            c = ws.cell(r, 2, text)
            c.alignment = Alignment(wrap_text=True, vertical="top")
            if kind == "h1":
                c.font = font(bold=True, size=16, color=C_HEAD)
            elif kind == "h2":
                c.font = font(bold=True, size=12, color=C_HEAD)
            elif kind == "sub":
                c.font = font(italic=True, color="595959")
            else:
                c.font = font()
            if text and kind == "":
                ws.row_dimensions[r].height = 15 * max(1, len(text) // 115 + 1)
            r += 1
        ws.protection.sheet = True

    # ----------------------------------------------------------------------------- sample data
    def make_sample(self) -> dict:
        rng = self.rng
        n = min(self.sample_days, len(self.days))
        days = self.days[:n]
        sample: dict = {}
        sample["prior_lti"] = date(self.year - 1, 9, 14)
        troy = float(self.s["constants"]["TROY_OZ_G"])
        stockpiles = [["ROM Pad", 20000, 1.2, "Y", date(self.year, 1, 1)], ["Stockpile HG", 5000, 2.1, "N", date(self.year, 1, 1)],
                      ["Stockpile MG", 30000, 1.0, "N", date(self.year, 1, 1)], ["Stockpile LG", 150000, 0.55, "N", date(self.year, 1, 1)],
                      ["Mineralised Waste Dump", 200000, 0.3, "N", date(self.year, 1, 1)]]
        sample["tblStockpiles"] = stockpiles
        # The plant draws from the Plant_Feed stockpile, so the synthetic head grade follows the ore actually
        # delivered there (direct tip plus the running ROM pad grade); otherwise the stockpile balance goes negative.
        feed_sp = next(r for r in stockpiles if r[3] == "Y")
        rom_t, rom_oz = float(feed_sp[1]), feed_sp[1] * feed_sp[2] / troy
        plant, moves = [], []
        for i, d in enumerate(days):
            hg_t, hg_g = round(rng.uniform(2200, 2800)), round(rng.uniform(2.0, 2.5), 2)
            tip_t, tip_g = round(rng.uniform(800, 1200)), round(rng.uniform(2.0, 2.5), 2)
            lg_t, lg_g = round(rng.uniform(1200, 1800)), round(rng.uniform(0.5, 0.7), 2)
            rh_t, rh_g = round(rng.uniform(2500, 3500)), 0.6
            moves.append([d, None, "Petowal Pit", "Ore HG", feed_sp[0], hg_t, hg_g, None, None, None])
            moves.append([d, None, "Petowal Pit", "Ore HG", "Crusher Direct Tip", tip_t, tip_g, None, None, None])
            moves.append([d, None, "Petowal Pit", "Ore LG", "Stockpile LG", lg_t, lg_g, None, None, None])
            moves.append([d, None, "Petowal Pit", "Waste", "Waste Dump", round(rng.uniform(16000, 20000)), None, None, None, None])
            moves.append([d, None, "Petowal Pit", "Mineralised Waste", "Mineralised Waste Dump", round(rng.uniform(600, 1000)), round(rng.uniform(0.3, 0.4), 2), None, None, None])
            moves.append([d, None, "Stockpile LG", "Ore LG", feed_sp[0], rh_t, rh_g, None, None, None])
            rom_t += hg_t + rh_t
            rom_oz += (hg_t * hg_g + rh_t * rh_g) / troy
            milled = round(rng.uniform(5800, 7000))
            draw = milled - tip_t
            head = round((tip_t * tip_g + draw * rom_oz / rom_t * troy) / milled * rng.uniform(0.97, 1.03), 2)
            rom_t -= draw
            rom_oz -= (milled * head - tip_t * tip_g) / troy
            tails = round(rng.uniform(0.08, 0.14), 3)
            planned = 8 if i % 14 == 13 else 0
            unplanned = round(rng.choice([0, 0, 0.5, 1, 2]), 1)
            run = round(24 - planned - unplanned - rng.uniform(0, 0.5), 1)
            poured = round(milled * (head - tails) / troy * 3 * rng.uniform(0.9, 1.1)) if i % 3 == 2 else 0
            plant.append([d, milled + rng.randint(-200, 300), milled, head, tails, round(rng.uniform(30, 60)), poured, run, planned, unplanned,
                          round(rng.uniform(14, 18), 1), round(milled * rng.uniform(0.35, 0.45)), round(milled * rng.uniform(1.2, 1.6)),
                          round(milled * rng.uniform(0.8, 1.0)), round(milled * rng.uniform(28, 34)), round(milled * rng.uniform(0.9, 1.2)),
                          round(rng.uniform(3000, 3500)), None])
        sample["tblPlant"] = plant
        sample["tblMovement"] = moves
        sample["tblMiningDaily"] = [[d, round(rng.uniform(1000, 1400)), round(rng.uniform(200, 400)), round(rng.uniform(22000, 28000)), round(rng.uniform(5000, 7000)),
                                     round(rng.uniform(30000, 40000)), round(rng.uniform(3000, 5000)), None] for d in days]
        fleet = []
        for d in days:
            fleet.append([d, "Excavator", 4, round(rng.uniform(6, 14), 1), round(rng.uniform(62, 78), 1), None])
            fleet.append([d, "Haul Truck", 20, round(rng.uniform(40, 80), 1), round(rng.uniform(300, 360), 1), None])
            fleet.append([d, "Drill", 3, round(rng.uniform(4, 10), 1), round(rng.uniform(40, 55), 1), None])
        sample["tblFleet"] = fleet
        safety = []
        for i, d in enumerate(days):
            safety.append([d, round(rng.uniform(3300, 3700)), round(rng.uniform(4300, 4700)), 0, 1 if i == 19 else 0, 0, 1 if i == 32 else 0,
                           rng.choice([0, 0, 0, 1]), rng.choice([0, 1, 2, 3]), rng.randint(8, 20), 1 if i == 40 else 0, 0, 1 if i == 25 else 0,
                           rng.choice([0, 0, 0, 1]), rng.randint(20, 40), rng.randint(2, 6), None])
        sample["tblSafety"] = safety
        sample["tblSafetyHistory"] = [[date(self.year - 1, m, 1), 240000 + m * 1000, 1 if m in (3, 9) else 0, 1 if m == 9 else 0] for m in range(1, 13)]
        gold = []
        for i, d in enumerate(days):
            if i % 10 == 9:
                gold.append([d, "Shipment", round(rng.uniform(55, 70), 1), round(rng.uniform(1100, 1400)), round(rng.uniform(300, 500)), f"SHP-{self.year}-{i // 10 + 1:03d}", None])
            if i % 10 == 4 and i > 10:
                gold.append([d, "Sale", None, round(rng.uniform(1100, 1400)), round(rng.uniform(300, 500)), f"SAL-{self.year}-{i // 10:03d}", None])
        sample["tblGold"] = gold
        sample["tblBudget"] = [[date(self.year, m, 1), 130000, 1.6, 600000, 195000, 1.55, 0.93, 9000, 250000, 36000, 1100000] for m in range(1, 13)]
        areas = self.s["lists"]["Area"]
        comments = []
        texts = ["Mill feed constrained by crusher liner change.", "Blast delayed to afternoon, fumes clearance.", "Two hazards closed out in the workshop.",
                 "TSF lift progressing, 40 percent complete.", "Community meeting held at Tomboronkoto.", "Excavator EX02 back in service after final drive replacement.",
                 "Grade control reconciliation within 3 percent for the week.", "Power interruption 40 minutes, no process upset.", "Cyanide delivery received, stock at 21 days."]
        for i, d in enumerate(days):
            for j in range(rng.randint(2, 3)):
                comments.append([d, areas[(i + j * 3) % len(areas)], texts[(i + j) % len(texts)], "DEMO"])
        sample["tblCommentary"] = comments
        return sample


def build(schema_path: Path, out: Path, year: int, sample_days: int = 0, report_date: date | None = None) -> Builder:
    schema = json.loads(Path(schema_path).read_text(encoding="utf-8"))
    b = Builder(schema, year, sample_days, report_date=report_date)
    wb = b.build()
    if sample_days:
        wb.properties.title += " DEMO DATA"
        ref = wb.defined_names["cfg_ReportTitle"].attr_text  # e.g. Config!$B$13
        wb["Config"][ref.split("!")[1].replace("$", "")] = f"{schema['site']['name']} daily production report  [DEMO DATA]"
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    return b


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--schema", default=str(DEFAULT_SCHEMA))
    ap.add_argument("--out", required=True)
    ap.add_argument("--year", type=int, default=2026)
    ap.add_argument("--sample-days", type=int, default=0, help="fill N days of synthetic DEMO data")
    ap.add_argument("--report-date", default=None, help="fix the Daily_Report date (YYYY-MM-DD) instead of the TODAY()-1 default; used by tests")
    a = ap.parse_args(argv)
    rd = date.fromisoformat(a.report_date) if a.report_date else None
    b = build(Path(a.schema), Path(a.out), a.year, a.sample_days, rd)
    print(f"wrote {a.out}: {len(b.wb.sheetnames)} sheets, {len(b.tables)} tables, {len(b.wb.defined_names)} names, ~{b.formula_count:,} formulas")


if __name__ == "__main__":
    main()
'@
Write-TextFile (Join-Path $Root "01_Tools\mps\build_workbook.py") $content_mps_build_workbook_py $false
$content_mps_build_ddl_py = @'
#!/usr/bin/env python3
"""
build_ddl.py - generate the SQL schema of the Mako Production System (MPS) web phase
from mps_schema.json, the same schema that drives build_workbook.py.

    python3 mps/build_ddl.py

writes, with no arguments and byte-identical output on every run:

    mps/sql/schema_postgres.sql   PostgreSQL 13 or later
    mps/sql/schema_sqlite.sql     SQLite 3.31 or later (generated columns)
    docs/data_dictionary.md       field-level dictionary and the Excel to SQL mapping

Everything about the data model comes from the schema and from the KPI catalogue
in build_workbook.py (KPIS): table and column names (schema keys sql_table,
sql_name), natural keys (unique), dialect-specific calc expressions (sql_postgres,
sql_sqlite) and the constants ({TROY}, {HOURS_PER_DAY}, {RATE_BASIS_HOURS} in calc
sql expressions). The only knowledge typed into this file is listed under MODEL
RULES below.

Mapping (documented again in the header of each SQL file and in the dictionary):
  * tables: Excel table name without the tbl prefix, CamelCase to snake_case;
    tables prefilled with one row per day get the suffix _daily
    (tblPlant -> plant_daily, tblMovement -> movement, tblMiningDaily -> mining_daily)
  * columns: the Excel field name in lower case (Milled_t -> milled_t)
  * date -> DATE, text -> TEXT, list -> TEXT + FOREIGN KEY, number -> NUMERIC(18,4),
    int -> INTEGER, pct -> NUMERIC(9,6) CHECK 0..1
  * lists -> ref_<list>(code TEXT PRIMARY KEY, sort_order) seeded from the schema;
    reference tables (Sources, Destinations, Stockpiles) -> real tables with seeds
  * calc fields with an "sql" expression -> GENERATED ALWAYS AS ... STORED (PostgreSQL),
    VIRTUAL (SQLite) or a <table>_calc view when a dialect cannot generate them;
    Excel lookups (INDEX/MATCH into a reference table) -> LEFT JOIN in <table>_calc
  * kpi_daily / kpi_monthly views reproduce the workbook KPI catalogue
"""
from __future__ import annotations

import json
import re
import sqlite3
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
SCHEMA_PATH = HERE / "schema" / "mps_schema.json"
OUT_POSTGRES = HERE / "sql" / "schema_postgres.sql"
OUT_SQLITE = HERE / "sql" / "schema_sqlite.sql"
OUT_DICTIONARY = ROOT / "docs" / "data_dictionary.md"

sys.path.insert(0, str(HERE))
try:
    from build_workbook import KPIS, CONSTANTS  # the KPI catalogue is the single source of KPI definitions
except ImportError as exc:  # pragma: no cover - environment problem, not a schema problem
    sys.exit(f"build_ddl.py needs build_workbook.py (and openpyxl) next to it: {exc}")

SQL_TYPES = {
    "date": "DATE",
    "text": "TEXT",
    "list": "TEXT",
    "number": "NUMERIC(18,4)",
    "int": "INTEGER",
    "pct": "NUMERIC(9,6)",
}
NUMERIC_TYPES = ("number", "pct")

# ----------------------------------------------------------------------------- MODEL RULES
# Rules the schema does not carry yet.

# Reference table keys that must also exist as keys of other reference tables
# (schema desc of Stockpiles.Stockpile: "must match a Source and a Destination").
REF_KEY_MATCHES = {"Stockpiles": ["Sources", "Destinations"]}

# KPI catalogue entries whose Excel Day formula is neither a plain SUMIFS over one table
# nor arithmetic over other KPIs. "agg" runs inside the per-table aggregate (GROUP BY date);
# {num:column} is the column cast to REAL in SQLite; {HOURS_PER_DAY} and the other schema constants are
# substituted. "custom" names a handler in Model.custom_kpi that runs in the final SELECT over the base
# row b (one row per date).
KPI_SQL = {
    # IF(COUNTIFS(run)+COUNTIFS(planned)+COUNTIFS(unplanned)>0,{HOURS_PER_DAY},0): a full day of calendar
    # hours for each day with any mill hours entered, else 0
    "Mill_Calendar_h": {"agg": ("tblPlant", "CASE WHEN COUNT(COALESCE(mill_run_h, mill_planned_maint_h, mill_unplanned_down_h)) > 0 THEN {HOURS_PER_DAY} ELSE 0 END")},
    # point value: the day's GIC when reported, else NULL (Excel shows blank)
    "GIC_oz": {"agg": ("tblPlant", "NULLIF(SUM({num:gic_oz}), 0)"), "nullable": True},
    # rolling 12 calendar months plus month to date, per million hours; prior-year months come from
    # safety_history for months that have no safety_daily rows (the workbook uses cfg_YearStart as the cut)
    "TRIFR_12m": {"custom": "rate12", "column": "recordables"},
    "LTIFR_12m": {"custom": "rate12", "column": "lti"},
    # days since the last LTI in safety_daily (the workbook also reads cfg_PriorLTIDate; the web system
    # should carry the last prior LTI into safety_daily or extend safety_history)
    "Days_Since_LTI": {"custom": "days_since_lti"},
}

SQL_WORDS = {
    "case", "when", "then", "else", "end", "and", "or", "not", "like", "is", "null", "in", "between", "cast", "as",
    "integer", "numeric", "real", "text", "interval", "extract", "day", "month", "year", "from", "true", "false",
    "timestamp", "date", "coalesce", "nullif", "abs", "round", "date_trunc", "strftime", "julianday",
}

LOOKUP_RE = re.compile(r"INDEX\((tbl\w+)\[(\w+)\],MATCH\(\{(\w+)\},\1\[(\w+)\],0\)\)")
SUMIFS_RE = re.compile(r"^SUMIFS\((tbl\w+)\[(\w+)\],\1\[Date\],\{d\}((?:,\1\[\w+\],(?:\"[^\"]*\"|-?[\d.]+))*)\)$")
CRIT_RE = re.compile(r",tbl\w+\[(\w+)\],(\"[^\"]*\"|-?[\d.]+)")
REF_RE = re.compile(r"\{([A-Za-z_][A-Za-z0-9_]*)\}")
IDENT_RE = re.compile(r"\b[A-Za-z_][A-Za-z0-9_]*\b")


# ----------------------------------------------------------------------------- naming
def snake(excel_table: str) -> str:
    n = re.sub(r"^tbl", "", excel_table)
    n = re.sub(r"(?<=[a-z0-9])(?=[A-Z])", "_", n)
    return n.lower()


def cname(field: dict) -> str:
    return field.get("sql_name") or field["name"].lower()


def sql_str(value) -> str:
    if value is None:
        return "NULL"
    if isinstance(value, (int, float)):
        return repr(value)
    return "'" + str(value).replace("'", "''") + "'"


def rewrite_idents(expr: str, mapping) -> str:
    """Replace bare identifiers outside string literals; mapping is a callable name -> replacement or None."""
    out = []
    for part in re.split(r"('(?:[^']|'')*')", expr):
        if part.startswith("'"):
            out.append(part)
        else:
            out.append(IDENT_RE.sub(lambda m: mapping(m.group(0)) or m.group(0), part))
    return "".join(out)


# ----------------------------------------------------------------------------- model
class Model:
    def __init__(self, schema: dict, sqlite_generated: bool = True):
        self.s = schema
        self.sqlite_generated = sqlite_generated
        self.troy = schema["constants"]["TROY_OZ_G"]
        self.rate_basis = schema["constants"].get("RATE_BASIS_HOURS", 1000000)
        # {NAME} placeholders in sql expressions and KPI_SQL, and the Config names used by the KPI catalogue
        self.consts = {"{TROY}": self.troy}
        for k, value in schema["constants"].items():
            self.consts["{" + k + "}"] = value
            if k in CONSTANTS:
                self.consts[CONSTANTS[k][0]] = value  # cfg_TroyOz and friends in the KPI catalogue
        self.lists = dict(schema["lists"])
        self.lists.update(schema.get("derived_lists", {}))
        self.ref_tables = schema["reference_tables"]
        self.by_excel = {t["table"]: t for t in schema["tables"]}
        for rname, rdef in self.ref_tables.items():
            self.by_excel[rdef["table"]] = self.ref_tdef(rname)
        # list name that is really a reference table key -> (sql table, sql column)
        self.ref_key_lists = {rdef["key"]: (self.tname(self.ref_tdef(r)), rdef["key"].lower()) for r, rdef in self.ref_tables.items()}
        self.notes: list[str] = []  # generation notes for the file headers

    # ---- schema access ----------------------------------------------------------
    def ref_tdef(self, rname: str) -> dict:
        rdef = self.ref_tables[rname]
        return {"table": rdef["table"], "sheet": rdef["sheet"], "title": rname, "fields": rdef["columns"], "key_field": rdef["key"],
                "rows": rdef["rows"], "ref_name": rname, "is_ref": True}

    def tname(self, tdef: dict) -> str:
        if tdef.get("sql_table"):
            return tdef["sql_table"]
        n = snake(tdef["table"])
        if tdef.get("prefill") == "dates" and not n.endswith("_daily"):
            n += "_daily"
        return n

    @staticmethod
    def input_fields(tdef: dict) -> list[dict]:
        return [f for f in tdef["fields"] if not f.get("calc")]

    @staticmethod
    def calc_fields(tdef: dict) -> list[dict]:
        return [f for f in tdef["fields"] if f.get("calc")]

    def key_field(self, tdef: dict) -> dict | None:
        """Field that is the primary key: the key field of prefilled tables, the key of reference tables."""
        if tdef.get("is_ref"):
            return next(f for f in tdef["fields"] if f["name"] == tdef["key_field"])
        if tdef.get("prefill"):
            return next((f for f in tdef["fields"] if f.get("key")), None)
        return None

    def natural_key(self, tdef: dict) -> list[str]:
        """Schema key "unique": a list of column groups; one group per table is supported."""
        groups = tdef.get("unique") or []
        if groups and not isinstance(groups[0], list):
            groups = [groups]
        if len(groups) > 1:
            raise NotImplementedError(f"{tdef['table']}: only one unique column group is supported")
        return list(groups[0]) if groups else []

    def const_sql(self, expr: str) -> str:
        """Replace {TROY}, {HOURS_PER_DAY}, ... and their cfg_ names with the schema constants."""
        for k, v in self.consts.items():
            expr = expr.replace(k, repr(v))
        return expr

    def list_target(self, f: dict) -> tuple[str, str]:
        lst = f["list"]
        if lst in self.ref_key_lists:
            return self.ref_key_lists[lst]
        if lst in self.lists:
            return f"ref_{lst.lower()}", "code"
        raise KeyError(f"field {f['name']} refers to unknown list {lst}")

    def lookup_of(self, f: dict) -> dict | None:
        m = LOOKUP_RE.search(f.get("calc", ""))
        if not m:
            return None
        ref_tdef = self.by_excel[m.group(1)]
        ref_col = next(c for c in ref_tdef["fields"] if c["name"] == m.group(2))
        return {"table": self.tname(ref_tdef), "column": cname(ref_col), "local": m.group(3).lower(), "key": m.group(4).lower()}

    # ---- calculated fields -------------------------------------------------------
    def calc_expr(self, tdef: dict, f: dict, dialect: str) -> str | None:
        expr = f.get(f"sql_{dialect}") or f.get("sql")
        if not expr:
            return None
        return self.const_sql(expr)

    def same_row_inputs_only(self, tdef: dict, expr: str) -> bool:
        body = re.sub(r"'(?:[^']|'')*'", "''", expr)
        body = re.sub(r"\b[A-Za-z_][A-Za-z0-9_]*\s*\(", "(", body)  # drop function names
        inputs = {cname(g) for g in self.input_fields(tdef)}
        for ident in set(IDENT_RE.findall(body)):
            if ident.lower() in SQL_WORDS or ident in inputs:
                continue
            return False
        return True

    def dialect_expr(self, tdef: dict, expr: str, dialect: str, prefix: str = "") -> str:
        """Qualify column references and, for SQLite, cast numeric columns to REAL so that a division
        never falls into integer arithmetic (SQLite stores integral NUMERIC values as INTEGER)."""
        cols = {cname(g): g for g in tdef["fields"]}

        def rep(name):
            g = cols.get(name)
            if g is None:
                return None
            ref = f"{prefix}{name}"
            if dialect == "sqlite" and g["type"] in NUMERIC_TYPES:
                return f"CAST({ref} AS REAL)"
            return ref if prefix else None

        return rewrite_idents(expr, rep)

    def sqlite_can_generate(self, tdef: dict, f: dict, expr: str) -> bool:
        con = sqlite3.connect(":memory:")
        cols = ", ".join(f"{cname(g)} {SQL_TYPES[g['type']]}" for g in self.input_fields(tdef))
        try:
            con.execute(f"CREATE TABLE t ({cols}, {cname(f)} {SQL_TYPES[f['type']]} GENERATED ALWAYS AS ({expr}) VIRTUAL)")
            con.execute("INSERT INTO t DEFAULT VALUES")
            con.execute("SELECT * FROM t").fetchall()
            return True
        except sqlite3.Error:
            return False
        finally:
            con.close()

    def plan_calcs(self, tdef: dict, dialect: str) -> dict:
        """Decide for each calc field: generated column, view column, lookup join or Excel-only."""
        plan = {"generated": [], "view": [], "lookup": [], "excel_only": []}
        for f in self.calc_fields(tdef):
            lk = self.lookup_of(f)
            if lk:
                plan["lookup"].append((f, lk))
                continue
            expr = self.calc_expr(tdef, f, dialect)
            if expr is None:
                plan["excel_only"].append(f)
                continue
            if self.same_row_inputs_only(tdef, expr):
                if dialect == "postgres":
                    plan["generated"].append((f, expr))
                    continue
                sq = self.dialect_expr(tdef, expr, "sqlite")
                if self.sqlite_generated and self.sqlite_can_generate(tdef, f, sq):
                    plan["generated"].append((f, sq))
                    continue
            plan["view"].append((f, expr))
        return plan

    # ---- DDL ---------------------------------------------------------------------
    def id_column(self, dialect: str) -> str:
        return "id BIGSERIAL PRIMARY KEY" if dialect == "postgres" else "id INTEGER PRIMARY KEY AUTOINCREMENT"

    def audit_columns(self) -> list[str]:
        return ["created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP",
                "updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP",
                "updated_by TEXT"]

    def first_of_month_check(self, col: str, dialect: str) -> str:
        if dialect == "postgres":
            return f"EXTRACT(DAY FROM {col}) = 1"
        return f"strftime('%d', {col}) = '01'"

    def check_for(self, f: dict) -> str | None:
        col = cname(f)
        t = f["type"]
        if t == "pct":
            lo = max(0, f.get("min", 0))
            hi = min(1, f.get("max", 1))
            return f"{col} BETWEEN {lo} AND {hi}"
        if t in ("number", "int") and ("min" in f or "max" in f):
            if "min" in f and "max" in f:
                return f"{col} BETWEEN {f['min']} AND {f['max']}"
            if "min" in f:
                return f"{col} >= {f['min']}"
            return f"{col} <= {f['max']}"
        return None

    def validation_text(self, tdef: dict, f: dict) -> str:
        if f.get("calc"):
            return ""
        if f["type"] == "list":
            tbl, col = self.list_target(f)
            return f"list {f['list']} ({tbl}.{col})"
        if f["type"] == "pct":
            return f"{max(0, f.get('min', 0))} to {min(1, f.get('max', 1))}"
        if "min" in f or "max" in f:
            return f"{f.get('min', '')} to {f.get('max', '')}".strip()
        if f["type"] == "date" and tdef.get("prefill") in ("months", "prior_months") and f.get("key"):
            return "first day of month"
        return ""

    def table_ddl(self, tdef: dict, dialect: str) -> list[str]:
        name = self.tname(tdef)
        key = self.key_field(tdef)
        cols: list[str] = []
        cons: list[str] = []
        if key is None:
            cols.append(self.id_column(dialect))
        for f in self.input_fields(tdef):
            col = cname(f)
            parts = [col, SQL_TYPES[f["type"]]]
            if f.get("required") or f.get("key") or f is key:
                parts.append("NOT NULL")
            cols.append(" ".join(parts))
            ck = self.check_for(f)
            if ck:
                cons.append(f"CONSTRAINT ck_{name}_{col} CHECK ({ck})")
            if f["type"] == "date" and f is key and tdef.get("prefill") in ("months", "prior_months"):
                cons.append(f"CONSTRAINT ck_{name}_{col}_first_of_month CHECK ({self.first_of_month_check(col, dialect)})")
            if f["type"] == "list":
                rt, rc = self.list_target(f)
                cons.append(f"CONSTRAINT fk_{name}_{col} FOREIGN KEY ({col}) REFERENCES {rt} ({rc})")
        plan = self.plan_calcs(tdef, dialect)
        for f, expr in plan["generated"]:
            storage = "STORED" if dialect == "postgres" else "VIRTUAL"
            cols.append(f"{cname(f)} {SQL_TYPES[f['type']]} GENERATED ALWAYS AS ({expr}) {storage}")
        cols += self.audit_columns()
        if key is not None:
            cons.insert(0, f"CONSTRAINT pk_{name} PRIMARY KEY ({cname(key)})")
        if tdef.get("is_ref"):
            for other in REF_KEY_MATCHES.get(tdef["ref_name"], []):
                ot = self.ref_tdef(other)
                cons.append(f"CONSTRAINT fk_{name}_{self.tname(ot)} FOREIGN KEY ({cname(key)}) REFERENCES {self.tname(ot)} ({self.ref_tables[other]['key'].lower()})")
        nk = self.natural_key(tdef)
        stmts: list[str] = []
        by_name = {f["name"]: f for f in tdef["fields"]}
        if nk:
            nk_fields = [by_name[n] for n in nk]
            nullable = [f for f in nk_fields if not (f.get("required") or f.get("key"))]
            if not nullable:
                cons.append(f"CONSTRAINT uq_{name}_key UNIQUE ({', '.join(cname(f) for f in nk_fields)})")
            else:
                parts = []
                for f in nk_fields:
                    if f in nullable:
                        if f["type"] not in ("text", "list"):
                            raise NotImplementedError(f"nullable non-text column {f['name']} in the natural key of {name}")
                        parts.append(f"COALESCE({cname(f)}, '')")
                    else:
                        parts.append(cname(f))
                stmts.append(f"CREATE UNIQUE INDEX IF NOT EXISTS ux_{name}_key ON {name} ({', '.join(parts)});")
        body = ",\n".join("    " + c for c in cols + cons)
        stmts.insert(0, f"CREATE TABLE IF NOT EXISTS {name} (\n{body}\n);")
        # index on date for tables where the date is neither the key nor the head of the natural key
        date_f = next((f for f in self.input_fields(tdef) if f["type"] == "date" and f["name"] == "Date"), None)
        if date_f is not None and key is None and (not nk or nk[0] != "Date"):
            stmts.append(f"CREATE INDEX IF NOT EXISTS ix_{name}_date ON {name} (date);")
        return stmts

    def comments_ddl(self, tdef: dict) -> list[str]:
        """PostgreSQL COMMENT ON statements from labels and descriptions."""
        name = self.tname(tdef)
        out = [f"COMMENT ON TABLE {name} IS {sql_str(tdef.get('title', name) + (' (Excel ' + tdef['table'] + ')'))};"]
        for f in tdef["fields"]:
            if self.lookup_of(f) or (f.get("calc") and not self.calc_expr(tdef, f, "postgres")):
                continue
            text = f.get("label", f["name"])
            if f.get("unit"):
                text += f" [{f['unit']}]"
            if f.get("desc"):
                text += ": " + f["desc"]
            out.append(f"COMMENT ON COLUMN {name}.{cname(f)} IS {sql_str(text)};")
        return out

    def calc_view(self, tdef: dict, dialect: str) -> list[str]:
        plan = self.plan_calcs(tdef, dialect)
        if not plan["lookup"] and not plan["view"]:
            return []
        name = self.tname(tdef)
        select = ["t.*"]
        joins = []
        for i, (f, lk) in enumerate(plan["lookup"], start=1):
            select.append(f"j{i}.{lk['column']} AS {cname(f)}")
            joins.append(f"LEFT JOIN {lk['table']} j{i} ON j{i}.{lk['key']} = t.{lk['local']}")
        for f, expr in plan["view"]:
            select.append(f"{self.dialect_expr(tdef, expr, dialect, prefix='t.')} AS {cname(f)}")
        sql = f"CREATE VIEW {name}_calc AS\nSELECT " + ",\n       ".join(select) + f"\nFROM {name} t"
        if joins:
            sql += "\n" + "\n".join(joins)
        return [f"DROP VIEW IF EXISTS {name}_calc;", sql + ";"]

    def has_calc_view(self, tdef: dict, dialect: str) -> bool:
        plan = self.plan_calcs(tdef, dialect)
        return bool(plan["lookup"] or plan["view"])

    def touch_trigger(self, tdef: dict, dialect: str) -> list[str]:
        name = self.tname(tdef)
        if dialect == "postgres":
            return [f"DROP TRIGGER IF EXISTS trg_{name}_touch ON {name};",
                    f"CREATE TRIGGER trg_{name}_touch BEFORE UPDATE ON {name} FOR EACH ROW EXECUTE FUNCTION mps_touch_updated_at();"]
        return [f"CREATE TRIGGER IF NOT EXISTS trg_{name}_touch AFTER UPDATE ON {name} FOR EACH ROW\n"
                f"WHEN NEW.updated_at IS OLD.updated_at\n"
                f"BEGIN\n    UPDATE {name} SET updated_at = CURRENT_TIMESTAMP WHERE rowid = NEW.rowid;\nEND;"]

    # ---- seeds -------------------------------------------------------------------
    def seed_lists(self, dialect: str) -> list[str]:
        out = []
        for lname, values in self.lists.items():
            rows = ", ".join(f"({sql_str(v)}, {i})" for i, v in enumerate(values, start=1))
            out.append(self.insert_ignore(f"ref_{lname.lower()}", "code, sort_order", rows, "code", dialect))
        return out

    def seed_refs(self, dialect: str) -> list[str]:
        out = []
        for rname in self.ref_tables:
            tdef = self.ref_tdef(rname)
            cols = ", ".join(cname(f) for f in tdef["fields"])
            rows = ", ".join("(" + ", ".join(sql_str(v) for v in row) + ")" for row in tdef["rows"])
            out.append(self.insert_ignore(self.tname(tdef), cols, rows, cname(self.key_field(tdef)), dialect))
        return out

    @staticmethod
    def insert_ignore(table: str, cols: str, rows: str, key: str, dialect: str) -> str:
        if dialect == "postgres":
            return f"INSERT INTO {table} ({cols}) VALUES {rows} ON CONFLICT ({key}) DO NOTHING;"
        return f"INSERT OR IGNORE INTO {table} ({cols}) VALUES {rows};"

    # ---- KPI views ---------------------------------------------------------------
    def kpi_source(self, excel_table: str, dialect: str) -> str:
        tdef = self.by_excel[excel_table]
        name = self.tname(tdef)
        return f"{name}_calc" if self.has_calc_view(tdef, dialect) else name

    def num(self, excel_table: str, col: str, dialect: str, prefix: str = "") -> str:
        """Column reference for aggregation, cast to REAL in SQLite when the source field is numeric."""
        tdef = self.by_excel[excel_table]
        f = next((g for g in tdef["fields"] if cname(g) == col), None)
        if dialect == "sqlite" and f is not None and f["type"] in NUMERIC_TYPES:
            return f"CAST({prefix}{col} AS REAL)"
        return f"{prefix}{col}"

    def kpi_plan(self, dialect: str) -> tuple[list[dict], list[str]]:
        """Translate the workbook KPI catalogue. Returns (entries, skipped keys)."""
        entries, skipped = [], []
        for k in KPIS:
            key = k["key"]
            e = {"key": key, "col": key.lower(), "label": k["label"], "unit": k["unit"], "kind": k["kind"], "nullable": False}
            ov = KPI_SQL.get(key)
            if ov and "agg" in ov:
                tbl, expr = ov["agg"]
                expr = re.sub(r"\{num:(\w+)\}", lambda m: self.num(tbl, m.group(1), dialect), expr)
                e.update(mode="agg", table=tbl, agg=self.const_sql(expr), nullable=ov.get("nullable", False))
            elif ov and "custom" in ov:
                e.update(mode="custom", custom=ov["custom"], column=ov.get("column"), nullable=True)
            elif k["kind"] == "sum":
                m = SUMIFS_RE.match(k["day"])
                if m:
                    tbl, col = m.group(1), m.group(2).lower()
                    crit = []
                    for c, v in CRIT_RE.findall(m.group(3)):
                        if v == '"<>"':  # Excel: not blank
                            crit.append(f"{c.lower()} IS NOT NULL")
                            continue
                        lit = sql_str(v[1:-1]) if v.startswith('"') else v
                        crit.append(f"{c.lower()} = {lit}")
                    ref = self.num(tbl, col, dialect)
                    agg = f"SUM({ref})" if not crit else f"SUM(CASE WHEN {' AND '.join(crit)} THEN {ref} ELSE 0 END)"
                    e.update(mode="agg", table=tbl, agg=agg)
                elif re.fullmatch(r"[\s\d.+\-*/()]*(\{\w+\}[\s\d.+\-*/()]*)+", k["day"]):
                    e.update(mode="expr", expr=k["day"])
                else:
                    skipped.append(key)
                    continue
            elif k["kind"] == "ratio":
                e.update(mode="ratio", num=self.const_sql(k["num"]), den=self.const_sql(k["den"]))
                e["nullable"] = True
            else:
                skipped.append(key)
                continue
            entries.append(e)
        return entries, skipped

    def kpi_daily_view(self, dialect: str) -> tuple[list[str], list[str]]:
        entries, skipped = self.kpi_plan(dialect)
        by_key = {e["key"]: e for e in entries}
        # per-table aggregates
        tables: dict[str, list[dict]] = {}
        for e in entries:
            if e["mode"] == "agg":
                tables.setdefault(e["table"], []).append(e)
        ctes = []
        spine = " UNION ".join(f"SELECT date FROM {self.tname(self.by_excel[t])}" for t in tables)
        ctes.append(f"days AS (\n    {spine}\n)")
        alias = {}
        for i, (tbl, es) in enumerate(tables.items(), start=1):
            a = f"a{i}"
            alias[tbl] = a
            cols = ",\n           ".join(f"{e['agg']} AS {e['col']}" for e in es)
            ctes.append(f"agg_{snake(tbl)} AS (\n    SELECT date,\n           {cols}\n    FROM {self.kpi_source(tbl, dialect)}\n    GROUP BY date\n)")
        base_cols = ["d.date"]
        for tbl, es in tables.items():
            for e in es:
                ref = f"{alias[tbl]}.{e['col']}"
                base_cols.append(f"{ref} AS {e['col']}" if e["nullable"] else f"COALESCE({ref}, 0) AS {e['col']}")
        joins = "\n".join(f"    LEFT JOIN agg_{snake(tbl)} {alias[tbl]} ON {alias[tbl]}.date = d.date" for tbl in tables)
        ctes.append("base AS (\n    SELECT " + ",\n           ".join(base_cols) + "\n    FROM days d\n" + joins + "\n)")

        def resolve(expr: str) -> str:
            def rep(m):
                e = by_key[m.group(1)]
                if e["mode"] == "agg":
                    return f"b.{e['col']}"
                if e["mode"] == "expr":
                    return "(" + resolve(e["expr"]) + ")"
                raise KeyError(f"KPI {m.group(1)} cannot be used inside another KPI definition")
            return REF_RE.sub(rep, expr)

        select = ["b.date"]
        for e in entries:
            if e["mode"] == "agg":
                select.append(f"b.{e['col']}")
            elif e["mode"] == "expr":
                select.append(f"{resolve(e['expr'])} AS {e['col']}")
            elif e["mode"] == "ratio":
                num, den = resolve(e["num"]), resolve(e["den"])
                select.append(f"CASE WHEN ({den}) > 0 THEN ({num}) / ({den}) END AS {e['col']}")
            else:
                select.append(f"{self.custom_kpi(e, dialect)} AS {e['col']}")
        sql = "CREATE VIEW kpi_daily AS\nWITH " + ",\n".join(ctes) + "\nSELECT " + ",\n       ".join(select) + "\nFROM base b;"
        return [sql], skipped

    def custom_kpi(self, e: dict, dialect: str) -> str:
        if dialect == "postgres":
            month_start = "date_trunc('month', b.date::timestamp)::date"
            win_start = "(date_trunc('month', b.date::timestamp) - INTERVAL '12 months')::date"
            month_of_s2 = "date_trunc('month', s2.date::timestamp)::date"
        else:
            month_start = "date(b.date, 'start of month')"
            win_start = "date(b.date, 'start of month', '-12 months')"
            month_of_s2 = "date(s2.date, 'start of month')"
        if e["custom"] == "rate12":
            col = e["column"]
            daily = self.kpi_source("tblSafety", dialect)  # safety_daily, or its _calc view when hours_total lives there
            history = self.kpi_source("tblSafetyHistory", dialect)

            def window(expr_daily: str, expr_hist: str) -> str:
                return (f"((SELECT COALESCE(SUM({expr_daily}), 0) FROM {daily} s WHERE s.date >= {win_start} AND s.date <= b.date)"
                        f" + (SELECT COALESCE(SUM({expr_hist}), 0) FROM {history} h WHERE h.month >= {win_start} AND h.month < {month_start}"
                        f" AND NOT EXISTS (SELECT 1 FROM {daily} s2 WHERE {month_of_s2} = h.month)))")
            events = window(f"s.{col}", f"h.{col}")
            hrs = window(self.num("tblSafety", "hours_total", dialect, prefix="s."),
                         self.num("tblSafetyHistory", "hours_total", dialect, prefix="h."))
            return f"CASE WHEN {hrs} > 0 THEN {events} * {self.rate_basis} / {hrs} END"
        if e["custom"] == "days_since_lti":
            last = f"(SELECT MAX(s.date) FROM {self.kpi_source('tblSafety', dialect)} s WHERE s.lti > 0 AND s.date <= b.date)"
            if dialect == "postgres":
                return f"b.date - {last}"
            return f"CAST(julianday(b.date) - julianday({last}) AS INTEGER)"
        raise KeyError(e["custom"])

    def kpi_monthly_view(self, dialect: str) -> list[str]:
        entries, _ = self.kpi_plan(dialect)
        by_key = {e["key"]: e for e in entries}
        month = "date_trunc('month', k.date)::date" if dialect == "postgres" else "date(k.date, 'start of month')"

        def resolve(expr: str) -> str:
            def rep(m):
                e = by_key[m.group(1)]
                if e["mode"] in ("agg", "expr"):
                    return f"SUM(k.{e['col']})"
                raise KeyError(m.group(1))
            return REF_RE.sub(rep, expr)

        select = [f"{month} AS month", "COUNT(*) AS days_reported"]
        for e in entries:
            if e["mode"] in ("agg", "expr") and not e["nullable"]:
                select.append(f"SUM(k.{e['col']}) AS {e['col']}")
            elif e["mode"] == "ratio":
                num, den = resolve(e["num"]), resolve(e["den"])
                select.append(f"CASE WHEN ({den}) > 0 THEN ({num}) / ({den}) END AS {e['col']}")
        sql = "CREATE VIEW kpi_monthly AS\nSELECT " + ",\n       ".join(select) + f"\nFROM kpi_daily k\nGROUP BY {month};"
        return [sql]

    # ---- whole script --------------------------------------------------------------
    def table_mapping(self) -> list[tuple[str, str, str, str]]:
        """(Excel table, sheet, SQL table, primary key) for input and reference tables."""
        rows = []
        for tdef in self.s["tables"]:
            key = self.key_field(tdef)
            rows.append((tdef["table"], tdef["sheet"], self.tname(tdef), cname(key) if key else "id"))
        for rname in self.ref_tables:
            tdef = self.ref_tdef(rname)
            rows.append((tdef["table"], tdef["sheet"], self.tname(tdef), cname(self.key_field(tdef))))
        return rows

    def excel_only_fields(self) -> list[tuple[str, str]]:
        out = []
        for tdef in self.s["tables"]:
            for f in self.calc_fields(tdef):
                if not self.lookup_of(f) and not self.calc_expr(tdef, f, "postgres"):
                    out.append((tdef["table"], f["name"]))
        return out

    def script(self, dialect: str) -> str:
        s = self.s
        site = s["site"]
        L: list[str] = []
        title = "PostgreSQL" if dialect == "postgres" else "SQLite"
        L.append(f"-- {s['system']} ({site['name']}, {site['company']}) schema for {title}")
        L.append(f"-- Generated by mps/build_ddl.py from mps/schema/mps_schema.json version {s['version']}. Do not edit; regenerate.")
        L.append("-- Safe to re-run on an existing database: tables use IF NOT EXISTS, views are dropped and recreated,")
        L.append("-- seeds are inserted only when missing. Structural changes to existing tables need a migration.")
        if dialect == "postgres":
            L.append("-- Run: psql -v ON_ERROR_STOP=1 -1 -f schema_postgres.sql  (single transaction; needs plpgsql, PostgreSQL 13 or later)")
        else:
            L.append("-- Run: sqlite3 mps.db < schema_sqlite.sql  (SQLite 3.31 or later; every connection must set PRAGMA foreign_keys = ON)")
        L.append("--")
        L.append("-- Naming: Excel table -> SQL table (primary key)")
        for excel, sheet, name, pk in self.table_mapping():
            L.append(f"--   {excel:<18} sheet {sheet:<18} -> {name} ({pk})")
        L.append("--   lists           -> ref_<list>(code, sort_order), one row per drop-down value")
        L.append("-- Columns are the Excel field names in lower case. Units: tonnes dry metric, grades g/t,")
        L.append(f"-- gold troy ounces ({self.troy} g), percentages stored as fractions 0 to 1.")
        L.append("-- Types: date DATE, text TEXT, list TEXT + FOREIGN KEY, number NUMERIC(18,4), int INTEGER, pct NUMERIC(9,6) CHECK 0..1.")
        L.append("-- CHECK constraints come from the schema min/max; NULL means not reported and always passes.")
        if dialect == "postgres":
            L.append("-- Calculated fields are GENERATED ALWAYS AS ... STORED; Excel lookups (source type, destination type)")
            L.append("-- are LEFT JOINs in the <table>_calc views.")
        else:
            L.append("-- Calculated fields are GENERATED ALWAYS AS ... VIRTUAL with numeric operands cast to REAL")
            L.append("-- (SQLite stores integral NUMERIC values as INTEGER and would otherwise divide as integers);")
            L.append("-- fields SQLite cannot generate and Excel lookups live in the <table>_calc views.")
        L.append("-- kpi_daily reproduces the workbook Calc_Daily day columns; kpi_monthly gives the month totals")
        L.append("-- with ratios recomputed from summed numerators and denominators.")
        eo = self.excel_only_fields()
        if eo:
            L.append("-- Excel-only helper fields not carried to SQL: " + ", ".join(f"{t}.{f}" for t, f in eo) + ".")
        L.append("")
        if dialect == "sqlite":
            L.append("PRAGMA foreign_keys = ON;")
            L.append("")
        # 1. list reference tables
        L.append("-- ---------------------------------------------------------------- reference lists (drop-down values)")
        for lname in self.lists:
            L.append(f"CREATE TABLE IF NOT EXISTS ref_{lname.lower()} (\n    code TEXT PRIMARY KEY,\n    sort_order INTEGER NOT NULL\n);")
        L.append("")
        # 2. reference tables
        L.append("-- ---------------------------------------------------------------- reference tables (Lists and Config sheets)")
        ref_tdefs = [self.ref_tdef(r) for r in self.ref_tables]
        ref_tdefs.sort(key=lambda t: 1 if t["ref_name"] in REF_KEY_MATCHES else 0)  # dependants after their targets
        for tdef in ref_tdefs:
            L.extend(self.table_ddl(tdef, dialect))
        L.append("")
        # 3. input tables
        L.append("-- ---------------------------------------------------------------- input tables")
        for tdef in self.s["tables"]:
            L.append(f"-- {tdef['title']} (owner: {tdef.get('owner', '')})")
            L.extend(self.table_ddl(tdef, dialect))
            L.append("")
        # 4. comments (PostgreSQL)
        if dialect == "postgres":
            L.append("-- ---------------------------------------------------------------- comments")
            for tdef in ref_tdefs + list(self.s["tables"]):
                L.extend(self.comments_ddl(tdef))
            L.append("")
        # 5. audit triggers
        L.append("-- ---------------------------------------------------------------- audit: keep updated_at current")
        if dialect == "postgres":
            L.append("CREATE OR REPLACE FUNCTION mps_touch_updated_at() RETURNS trigger AS $$\nBEGIN\n    NEW.updated_at := CURRENT_TIMESTAMP;\n    RETURN NEW;\nEND;\n$$ LANGUAGE plpgsql;")
        for tdef in ref_tdefs + list(self.s["tables"]):
            L.extend(self.touch_trigger(tdef, dialect))
        L.append("")
        # 6. calc views
        L.append("-- ---------------------------------------------------------------- calculated views")
        L.append("DROP VIEW IF EXISTS kpi_monthly;")
        L.append("DROP VIEW IF EXISTS kpi_daily;")
        for tdef in self.s["tables"]:
            L.extend(self.calc_view(tdef, dialect))
        L.append("")
        # 7. KPI views
        L.append("-- ---------------------------------------------------------------- KPI views (workbook KPI catalogue)")
        L.append("-- Ex-pit = source type Pit; ore = material starting with Ore; rehandle = source type Stockpile.")
        L.append("-- Ratios are NULL when the denominator is zero. Sums are 0 on days without rows, as in the workbook.")
        stmts, skipped = self.kpi_daily_view(dialect)
        if skipped:
            L.append("-- KPI catalogue entries without an SQL translation: " + ", ".join(skipped))
        L.extend(stmts)
        L.extend(self.kpi_monthly_view(dialect))
        L.append("")
        # 8. seeds
        L.append("-- ---------------------------------------------------------------- seed data (inserted only when missing)")
        L.extend(self.seed_lists(dialect))
        L.extend(self.seed_refs(dialect))
        L.append("")
        return "\n".join(L)

    # ---- data dictionary -----------------------------------------------------------
    def dictionary(self) -> str:
        s = self.s
        site = s["site"]
        L: list[str] = []
        esc = lambda v: str(v).replace("|", "\\|").replace("\n", " ")  # noqa: E731
        L.append(f"# {s['system']} data dictionary")
        L.append("")
        L.append(f"{site['name']}, {site['company']} (owner {site['owner']}), {site['country']}. "
                 f"Generated from `mps/schema/mps_schema.json` version {s['version']} by `mps/build_ddl.py`; "
                 "the same schema generates the Excel workbook (`mps/build_workbook.py`). Do not edit this file; change the schema and regenerate.")
        L.append("")
        L.append(f"Conventions: metric units, tonnes dry, grades in g/t, gold in troy ounces ({self.troy} g), "
                 "percentages stored as fractions between 0 and 1, dates are production days (day shift start to night shift end), "
                 "site time zone " + site["timezone"] + ". Blank (NULL) means not reported; zero means reported as zero.")
        L.append("")
        L.append("## Contents")
        L.append("")
        L.append("1. [Excel to SQL mapping](#excel-to-sql-mapping)")
        L.append("2. [Input tables](#input-tables)")
        L.append("3. [Reference tables](#reference-tables)")
        L.append("4. [Lists](#lists)")
        L.append("5. [KPI catalogue](#kpi-catalogue)")
        L.append("")
        # mapping
        L.append("## Excel to SQL mapping")
        L.append("")
        L.append("| Excel table | Sheet | SQL table | Primary key | Rows | Owner |")
        L.append("|---|---|---|---|---|---|")
        for tdef in s["tables"]:
            key = self.key_field(tdef)
            rows = {"dates": "one per day", "months": "one per month", "prior_months": "one per prior-year month"}.get(tdef.get("prefill"), "many per day")
            L.append(f"| {tdef['table']} | {tdef['sheet']} | `{self.tname(tdef)}` | `{cname(key) if key else 'id'}` | {rows} | {esc(tdef.get('owner', ''))} |")
        for rname in self.ref_tables:
            tdef = self.ref_tdef(rname)
            L.append(f"| {tdef['table']} | {tdef['sheet']} | `{self.tname(tdef)}` | `{cname(self.key_field(tdef))}` | reference | Configuration |")
        L.append("")
        L.append("| Schema type | Excel | PostgreSQL and SQLite |")
        L.append("|---|---|---|")
        L.append("| date | date cell, dd-mmm-yyyy | `DATE` (SQLite: ISO text YYYY-MM-DD) |")
        L.append("| text | free text | `TEXT` |")
        L.append("| list | drop-down from the Lists sheet | `TEXT` with a `FOREIGN KEY` to `ref_<list>(code)` or to the reference table |")
        L.append("| number | decimal, min to max warning | `NUMERIC(18,4)` with `CHECK (min to max)` |")
        L.append("| int | whole number, min to max warning | `INTEGER` with `CHECK (min to max)` |")
        L.append("| pct | percentage, stored as a fraction | `NUMERIC(9,6)` with `CHECK (0 to 1)` |")
        L.append("")
        L.append("Rules applied by the generator:")
        L.append("")
        L.append("- Table names drop the `tbl` prefix and become snake_case; tables with one row per day get the suffix `_daily`. Column names are the Excel field names in lower case.")
        L.append("- Primary key: the date or month column of the tables prefilled with one row per period; an `id` column (`BIGSERIAL` in PostgreSQL, `INTEGER PRIMARY KEY AUTOINCREMENT` in SQLite) elsewhere. "
                 "Natural keys: " + "; ".join(f"`{self.tname(t)}` ({', '.join(c.lower() for c in self.natural_key(t))})" for t in self.s["tables"] if self.natural_key(t)) +
                 ". Where a natural-key column is optional (movement shift, blank for a daily total) the uniqueness is a unique index on `COALESCE(column, '')` so that two daily totals cannot be entered twice.")
        L.append("- `NOT NULL` for required and key fields. `CHECK` constraints from the schema min and max; Excel only warns, SQL rejects. Month columns must be the first day of the month.")
        L.append("- List fields reference `ref_<list>` tables seeded with the drop-down values in order (`sort_order`). `movement.source` and `movement.destination` reference the `sources` and `destinations` tables; `stockpiles.stockpile` must exist in both.")
        L.append("- Calculated fields with an SQL expression are generated columns (`GENERATED ALWAYS AS ... STORED` in PostgreSQL, `VIRTUAL` in SQLite, where numeric operands are cast to `REAL` to avoid integer division). "
                 "Lookups into reference tables (source type, destination type) are `LEFT JOIN`s in the `<table>_calc` views, and SQLite also puts there anything it cannot generate. "
                 "Generated columns carry no CHECK constraint: validation applies to inputs only.")
        eo = self.excel_only_fields()
        if eo:
            L.append("- Excel-only report helpers are not carried to SQL: " + ", ".join(f"`{t}.{f}`" for t, f in eo) + ".")
        L.append("- Every table except the `ref_*` lists carries `created_at`, `updated_at` (kept current by a trigger) and `updated_by TEXT`.")
        L.append("- `kpi_daily` reproduces the workbook Calc_Daily day columns from the KPI catalogue in `build_workbook.py`; `kpi_monthly` groups it by month with ratios recomputed from the summed numerators and denominators. Both views are recreated by every run of the script.")
        L.append("- Both scripts are idempotent (`IF NOT EXISTS`, seeds inserted only when missing, views recreated). Structural changes to a populated database need a migration.")
        L.append("")
        # input tables
        L.append("## Input tables")
        L.append("")
        for tdef in s["tables"]:
            L.extend(self.dict_table(tdef))
        L.append("## Reference tables")
        L.append("")
        for rname in self.ref_tables:
            tdef = self.ref_tdef(rname)
            L.extend(self.dict_table(tdef))
            cols = [cname(f) for f in tdef["fields"]]
            L.append("Seed rows:")
            L.append("")
            L.append("| " + " | ".join(f"`{c}`" for c in cols) + " |")
            L.append("|" + "---|" * len(cols))
            for row in tdef["rows"]:
                L.append("| " + " | ".join("" if v is None else esc(v) for v in row) + " |")
            L.append("")
        # lists
        L.append("## Lists")
        L.append("")
        L.append("Each list is a table `ref_<list>` with `code TEXT PRIMARY KEY` and `sort_order INTEGER`. Add a value by inserting a row; do not rename codes that are already referenced.")
        L.append("")
        L.append("| List | SQL table | Values (in order) | Used by |")
        L.append("|---|---|---|---|")
        for lname, values in self.lists.items():
            users = []
            for tdef in list(s["tables"]) + [self.ref_tdef(r) for r in self.ref_tables]:
                for f in tdef["fields"]:
                    if f.get("list") == lname:
                        users.append(f"`{self.tname(tdef)}.{cname(f)}`")
            L.append(f"| {lname} | `ref_{lname.lower()}` | {', '.join(values)} | {', '.join(users)} |")
        L.append("")
        # KPI catalogue
        L.append("## KPI catalogue")
        L.append("")
        L.append("Columns of the `kpi_daily` view (one row per date with data), in the order of the workbook KPI catalogue. "
                 "Sum KPIs are day totals (0 when no rows); ratio KPIs are numerator over denominator for the day and NULL when the denominator is 0; "
                 "`kpi_monthly` recomputes ratios from monthly sums, which is how the workbook builds month-to-date figures. "
                 "Ex-pit means source type Pit, ore means a material starting with Ore, rehandle means source type Stockpile.")
        L.append("")
        L.append("| Column | Label | Unit | Kind | Definition (PostgreSQL) |")
        L.append("|---|---|---|---|---|")
        entries, skipped = self.kpi_plan("postgres")
        by_key = {e["key"]: e for e in entries}

        def show(expr: str) -> str:
            return REF_RE.sub(lambda m: by_key[m.group(1)]["col"], expr)

        def term(expr: str) -> str:  # parenthesise compound operands so the precedence reads as in the view
            out = show(expr)
            return f"({out})" if re.search(r"[-+*/]", out) else out
        for e in entries:
            if e["mode"] == "agg":
                d = f"{e['agg']} over `{self.tname(self.by_excel[e['table']])}` per date"
            elif e["mode"] == "expr":
                d = show(e["expr"])
            elif e["mode"] == "ratio":
                d = f"{term(e['num'])} / {term(e['den'])}"
            else:
                d = {"rate12": f"({e.get('column')} over the last 12 calendar months plus month to date) * {self.rate_basis} / hours_total over the same window; prior months from safety_history where safety_daily has no rows",
                     "days_since_lti": "date minus the last date in safety_daily with lti > 0"}[e["custom"]]
            L.append(f"| `{e['col']}` | {esc(e['label'])} | {e['unit']} | {e['kind']} | {esc(d)} |")
        if skipped:
            L.append("")
            L.append("Not translated: " + ", ".join(skipped) + ".")
        L.append("")
        return "\n".join(L)

    def dict_table(self, tdef: dict) -> list[str]:
        L: list[str] = []
        esc = lambda v: str(v).replace("|", "\\|").replace("\n", " ")  # noqa: E731
        name = self.tname(tdef)
        key = self.key_field(tdef)
        L.append(f"### {name} ({tdef['table']}, sheet {tdef['sheet']})")
        L.append("")
        head = tdef.get("title", tdef.get("ref_name", name))
        if tdef.get("owner"):
            head += f". Owner: {tdef['owner']}"
        head += f". Primary key: `{cname(key) if key else 'id'}`"
        nk = self.natural_key(tdef)
        if nk:
            head += f"; unique: ({', '.join(c.lower() for c in nk)})"
        for other in REF_KEY_MATCHES.get(tdef.get("ref_name", ""), []):
            head += f"; `{cname(key)}` must exist in `{self.tname(self.ref_tdef(other))}`"
        plan = self.plan_calcs(tdef, "postgres")
        plan_lite = self.plan_calcs(tdef, "sqlite")
        if plan["lookup"] or plan["view"] or plan_lite["view"]:
            head += f". View `{name}_calc` adds the lookup columns" + (" (SQLite: and the fields it cannot generate)" if plan_lite["view"] else "")
        L.append(head + ".")
        L.append("")
        L.append("| Column | Excel field | Label | Type | Unit | Required | Validation | Description | Calculation |")
        L.append("|---|---|---|---|---|---|---|---|---|")
        gen = {f["name"]: expr for f, expr in plan["generated"]}
        view = {f["name"]: expr for f, expr in plan["view"]}
        look = {f["name"]: lk for f, lk in plan["lookup"]}
        for f in tdef["fields"]:
            calc = ""
            typ = SQL_TYPES[f["type"]]
            if f["name"] in gen:
                calc = f"`{esc(gen[f['name']])}`"
                typ = f"generated {typ}"
            elif f["name"] in view:
                calc = f"`{esc(view[f['name']])}` (view)"
                typ = f"view {typ}"
            elif f["name"] in look:
                lk = look[f["name"]]
                calc = f"lookup `{lk['table']}.{lk['column']}` (view)"
                typ = f"view {typ}"
            elif f.get("calc"):
                calc = f"Excel only: `{esc(f['calc'])}`"
                typ = "not in SQL"
            req = "key" if (f is key) else ("yes" if f.get("required") else "")
            L.append(f"| `{cname(f)}` | {f['name']} | {esc(f.get('label', f['name']))} | {f['type']} ({typ}) | {f.get('unit', '')} | {req} | "
                     f"{esc(self.validation_text(tdef, f))} | {esc(f.get('desc', ''))} | {calc} |")
        L.append("| `created_at` | | Created at | TIMESTAMP | | yes | default now | Audit: row creation time | |")
        L.append("| `updated_at` | | Updated at | TIMESTAMP | | yes | default now | Audit: last change, kept current by trigger | |")
        L.append("| `updated_by` | | Updated by | TEXT | | | | Audit: user who made the last change | |")
        L.append("")
        return L


# ----------------------------------------------------------------------------- entry point
def build(schema_path: Path = SCHEMA_PATH) -> dict[str, str]:
    schema = json.loads(Path(schema_path).read_text(encoding="utf-8"))
    m = Model(schema)
    return {
        str(OUT_POSTGRES): m.script("postgres"),
        str(OUT_SQLITE): m.script("sqlite"),
        str(OUT_DICTIONARY): m.dictionary(),
    }


def main(argv=None) -> int:
    outputs = build()
    for path, text in outputs.items():
        p = Path(path)
        p.parent.mkdir(parents=True, exist_ok=True)
        p.write_text(text + ("\n" if not text.endswith("\n") else ""), encoding="utf-8", newline="\n")
        print(f"wrote {p.relative_to(ROOT)} ({len(text.splitlines()):,} lines)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
'@
Write-TextFile (Join-Path $Root "01_Tools\mps\build_ddl.py") $content_mps_build_ddl_py $false
$content_mps_validate_workbook_py = @'
#!/usr/bin/env python3
"""
validate_workbook.py - build, recalculate and independently verify the MPS workbook.

What it does
  1. builds four workbooks with build_workbook.py:
       demo    : 74 days of synthetic DEMO data, report date fixed to 10 Feb
       default : the same data, report date left to the workbook's default formula
       blank   : no data at all
       edge    : the demo with partially entered rows (missing grades, missing run
                 hours, a full-day shutdown, an ungraded ore row, an unknown source,
                 a row dated in the prior year) plus a Probe sheet that measures every
                 lst_ drop-down name
  2. recalculates them with LibreOffice (headless) and reads the values back
  3. recomputes the key figures in plain Python from Builder.sample (the synthetic
     rows, in schema field order) without using any workbook formula, and compares
     them with Calc_Daily for 1 Jan, 10 Feb and the last day with data
  4. checks Daily_Report, Monthly_Summary and Chart_Data against Calc_Daily and
     against the independent figures
  5. scans every sheet for error values and audits the built file with
     tools/xl_inspect.py (no external links, no VBA, no dangling references,
     no broken names)
  6. checks that the blank workbook recalculates without a single error value

Usage (from the repository root):
    python3 mps/validate_workbook.py [--keep] [--workdir DIR] [--soffice PATH]

Prints one line per check (check, expected, actual, pass) and exits with code 1
when any check fails. Temporary files go to a fresh temp folder unless --workdir
is given; --keep leaves them in place.
"""
from __future__ import annotations

import argparse
import calendar
import json
import shutil
import subprocess
import sys
import tempfile
import time
import warnings
from datetime import date, datetime, time as dtime, timedelta
from pathlib import Path

import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.utils.datetime import from_excel

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

HERE = Path(__file__).resolve().parent
REPO = HERE.parent
sys.path.insert(0, str(HERE))
import build_workbook as bw  # noqa: E402

YEAR = 2026
SAMPLE_DAYS = 74
REPORT_DATE = date(YEAR, 2, 10)
CHECK_DATES = [date(YEAR, 1, 1), REPORT_DATE, date(YEAR, 1, 1) + timedelta(days=SAMPLE_DAYS - 1)]
EDGE_DATES = [date(YEAR, 2, d) for d in (5, 6, 7, 8, 9, 10)]


def set_year(year: int):
    """Point every date expectation at another reporting year (--year 2028 exercises the leap-year branch)."""
    global YEAR, REPORT_DATE, CHECK_DATES, EDGE_DATES
    YEAR = year
    REPORT_DATE = date(YEAR, 2, 10)
    CHECK_DATES = [date(YEAR, 1, 1), REPORT_DATE, date(YEAR, 1, 1) + timedelta(days=SAMPLE_DAYS - 1)]
    EDGE_DATES = [date(YEAR, 2, d) for d in (5, 6, 7, 8, 9, 10)]
TOL_SUM = 1e-6      # relative tolerance for additive figures (tonnes, ounces, hours, counts)
TOL_RATIO = 1e-4    # relative tolerance for ratios, grades, rates and pro-rata budgets
TOL_SAME = 1e-9     # two cells of the same recalculated workbook that must agree
ERROR_VALUES = ("#REF!", "#NAME?", "#VALUE!", "#DIV/0!", "#NUM!", "#NULL!", "#N/A")
MONTHS = [calendar.month_abbr[m] for m in range(1, 13)]


# ----------------------------------------------------------------------------- value helpers
def as_date(v):
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, (int, float)) and not isinstance(v, bool) and v > 20000:
        return from_excel(v).date()
    return None


def norm(v):
    """Normalise a cell value: blank strings become None, datetimes become dates, ints become floats."""
    if v is None:
        return None
    if isinstance(v, str):
        return None if v == "" else v
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, bool):
        return v
    if isinstance(v, (int, float)):
        return float(v)
    return v


def is_num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def num(v) -> float:
    """Excel N(): blank -> 0."""
    return float(v) if is_num(v) else 0.0


def close(exp, act, rel) -> bool:
    exp, act = norm(exp), norm(act)
    if exp is None or act is None:
        return exp is None and act is None
    if is_num(exp) and is_num(act):
        return abs(act - exp) <= rel * max(abs(exp), 1e-9)
    return exp == act


def div(n, d):
    """Ratio as the workbook defines it: blank unless the denominator is positive."""
    return n / d if d > 0 else None


def fmt(v) -> str:
    v = norm(v)
    if v is None:
        return "blank"
    if isinstance(v, float):
        if abs(v) < 1e15 and v == int(v):
            return str(int(v))
        return f"{v:.10g}"
    if isinstance(v, date):
        return v.isoformat()
    if isinstance(v, dtime):
        return "0 (zero date)"
    s = str(v)
    return s if len(s) <= 30 else s[:27] + "..."


class Checks:
    def __init__(self):
        self.rows: list[tuple[str, str, str, bool]] = []

    def add(self, name, expected, actual, ok):
        self.rows.append((name, str(expected), str(actual), bool(ok)))

    def num(self, name, expected, actual, rel=TOL_SUM):
        self.add(name, fmt(expected), fmt(actual), close(expected, actual, rel))

    def eq(self, name, expected, actual):
        self.add(name, fmt(expected), fmt(actual), norm(expected) == norm(actual))

    def group(self, name, total, mismatches):
        exp = f"{total} agree"
        act = exp if not mismatches else f"{len(mismatches)} differ: " + "; ".join(mismatches[:2])
        self.add(name, exp, act, not mismatches)

    @property
    def failed(self):
        return [r for r in self.rows if not r[3]]

    def print(self):
        w0 = min(max(len(r[0]) for r in self.rows), 64)
        w1 = min(max(len(r[1]) for r in self.rows), 30)
        print(f"{'check':<{w0}}  {'expected':<{w1}}  {'actual':<{w1}}  pass")
        print("-" * (w0 + 2 * w1 + 10))
        for name, exp, act, ok in self.rows:
            a = act if ok or len(act) <= 80 else act[:77] + "..."
            print(f"{name[:w0]:<{w0}}  {exp[:w1]:<{w1}}  {a:<{w1}}  {'ok' if ok else 'FAIL'}")
        print("-" * (w0 + 2 * w1 + 10))
        print(f"{len(self.rows)} checks, {len(self.failed)} failed")


# ----------------------------------------------------------------------------- sample access
class Sample:
    """Builder.sample rows (list of lists in schema field order) as dicts keyed by field name."""

    def __init__(self, builder, sample: dict | None = None):
        self.b = builder
        self.sample = sample if sample is not None else builder.sample
        s = builder.s
        self.troy = float(s["constants"]["TROY_OZ_G"])
        self.fields = {t["table"]: [f["name"] for f in t["fields"] if not f.get("calc")] for t in s["tables"]}
        for rdef in s["reference_tables"].values():
            self.fields[rdef["table"]] = [c["name"] for c in rdef["columns"]]
        self.src_type = {r[0]: r[1] for r in s["reference_tables"]["Sources"]["rows"]}
        self.dst_type = {r[0]: r[1] for r in s["reference_tables"]["Destinations"]["rows"]}
        self.width_problems = []
        self._cache = {}

    def rows(self, table) -> list[dict]:
        if table in self._cache:
            return self._cache[table]
        names = self.fields[table]
        out = []
        for row in self.sample.get(table, []):
            if len(row) != len(names):
                self.width_problems.append(f"{table}: row has {len(row)} values for {len(names)} fields")
            out.append({n: (row[i] if i < len(row) else None) for i, n in enumerate(names)})
        self._cache[table] = out
        return out

    def is_ore(self, material) -> bool:
        return str(material or "").startswith("Ore")

    def oz(self, r) -> float:
        """Contained ounces of a movement row: zero when tonnes or grade is blank."""
        if r.get("Tonnes_t") is None or r.get("Grade_gpt") is None:
            return 0.0
        return float(r["Tonnes_t"]) * float(r["Grade_gpt"]) / self.troy

    def last_plant_date(self):
        ds = [r["Date"] for r in self.rows("tblPlant") if r["Milled_t"] is not None]
        return max(ds) if ds else None


def in_period(rd: date, d: date, period: str) -> bool:
    if period == "Day":
        return rd == d
    if period == "MTD":
        return rd.year == d.year and rd.month == d.month and rd <= d
    return rd.year == d.year and rd <= d


def total(rows, d, period, fn) -> float:
    return float(sum(fn(r) for r in rows if in_period(r["Date"], d, period)))


# ----------------------------------------------------------------------------- independent expectations
# key -> True when additive (sum tolerance), False when a ratio (ratio tolerance)
KIND = {
    "Ore_Mined_t": True, "Ore_Mined_oz": True, "Ore_Grade_gpt": False, "Waste_t": True, "TMM_t": True, "Strip_Ratio": False,
    "Rehandle_t": True, "Milled_t": True, "Feed_oz": True, "Recovered_oz": True, "Head_Grade_gpt": False, "Recovery_pct": False,
    "Gold_Poured_oz": True, "Throughput_tph": False, "Mill_Availability_pct": False, "Mill_Utilisation_pct": False,
    "Cyanide_kgpt": False, "Hours_Worked": True, "Recordables": True, "LTI": True, "Gold_Shipped_oz": True, "Gold_Sold_oz": True,
    "GIC_oz": True, "Days_Since_LTI": True, "TRIFR_12m": False, "LTIFR_12m": False,
    "Ore_Ungraded_t": True, "Unclassified_t": True, "Tails_Grade_gpt": False, "Mill_Calendar_h": True,
    "GC_Drill_m": True, "Gravity_Share_pct": False, "Exc_Productivity_tph": False, "Trk_Productivity_tph": False,
    "Bud_Milled_t_MTD": False, "Bud_Milled_t_YTD": False, "Bud_Ore_Grade_gpt_MTD": False, "Bud_Recovery_pct_YTD": False,
}


def expected_kpis(S: Sample, d: date) -> dict[str, dict[str, object]]:
    """Recompute the KPI set for one date, per period, straight from the sample rows.
    Blank means not reported: a row whose grade, tails grade or run hours is blank contributes nothing to the
    ratio that needs it (the tonnes are paired with the assay or the hours that carry them)."""
    T = S.troy
    mv, pl, sf, gd = S.rows("tblMovement"), S.rows("tblPlant"), S.rows("tblSafety"), S.rows("tblGold")
    md, fl = S.rows("tblMiningDaily"), S.rows("tblFleet")

    def pit(r):
        return S.src_type.get(r["Source"]) == "Pit"

    def pit_ore(r):
        return pit(r) and S.is_ore(r["Material"])

    def graded(r):
        return r.get("Grade_gpt") is not None

    def milled(r):
        return num(r["Milled_t"])

    def has(r, *cols):
        return all(r[c] is not None for c in cols)

    def feed_oz(r):
        return milled(r) * num(r["Head_Grade_gpt"]) / T if has(r, "Milled_t", "Head_Grade_gpt") else 0.0

    def tails_oz(r):
        return milled(r) * num(r["Tails_Grade_gpt"]) / T if has(r, "Milled_t", "Tails_Grade_gpt") else 0.0

    def rec_oz(r):
        return feed_oz(r) - tails_oz(r) if has(r, "Milled_t", "Head_Grade_gpt", "Tails_Grade_gpt") else 0.0

    def hours(r):
        return num(r["Hours_Employees"]) + num(r["Hours_Contractors"])

    def recordables(r):
        return num(r["Fatality"]) + num(r["LTI"]) + num(r["RWI"]) + num(r["MTI"])

    out: dict[str, dict[str, object]] = {}

    def put(key, period, val):
        out.setdefault(key, {})[period] = val

    for p in ("Day", "MTD", "YTD"):
        ore_t = total(mv, d, p, lambda r: float(r["Tonnes_t"]) if pit_ore(r) else 0.0)
        ore_graded_t = total(mv, d, p, lambda r: float(r["Tonnes_t"]) if pit_ore(r) and graded(r) else 0.0)
        ore_oz = total(mv, d, p, lambda r: S.oz(r) if pit_ore(r) else 0.0)
        waste_t = total(mv, d, p, lambda r: float(r["Tonnes_t"]) if pit(r) and not S.is_ore(r["Material"]) else 0.0)
        rehandle = total(mv, d, p, lambda r: float(r["Tonnes_t"]) if S.src_type.get(r["Source"]) == "Stockpile" else 0.0)
        moved = total(mv, d, p, lambda r: float(r["Tonnes_t"]))
        mil = total(pl, d, p, milled)
        mil_head = total(pl, d, p, lambda r: milled(r) if r["Head_Grade_gpt"] is not None else 0.0)
        mil_tails = total(pl, d, p, lambda r: milled(r) if r["Tails_Grade_gpt"] is not None else 0.0)
        mil_timed = total(pl, d, p, lambda r: milled(r) if r["Mill_Run_h"] is not None else 0.0)
        fz = total(pl, d, p, feed_oz)
        fz_recon = total(pl, d, p, lambda r: feed_oz(r) if r["Tails_Grade_gpt"] is not None else 0.0)
        tz = total(pl, d, p, tails_oz)
        rz = total(pl, d, p, rec_oz)
        poured = total(pl, d, p, lambda r: num(r["Gold_Poured_oz"]))
        run_h = total(pl, d, p, lambda r: num(r["Mill_Run_h"]))
        plan_h = total(pl, d, p, lambda r: num(r["Mill_Planned_Maint_h"]))
        unpl_h = total(pl, d, p, lambda r: num(r["Mill_Unplanned_Down_h"]))
        cal_h = total(pl, d, p, lambda r: 24.0 if any(r[c] is not None for c in ("Mill_Run_h", "Mill_Planned_Maint_h", "Mill_Unplanned_Down_h")) else 0.0)
        cn_kg = total(pl, d, p, lambda r: num(r["Cyanide_kg"]))
        hrs = total(sf, d, p, hours)
        rec = total(sf, d, p, recordables)
        lti = total(sf, d, p, lambda r: num(r["LTI"]))
        shipped = total(gd, d, p, lambda r: num(r["Gold_oz"]) if r["Type"] == "Shipment" else 0.0)
        gravity = total(pl, d, p, lambda r: num(r["Gravity_Gold_oz"]))
        gc_m = total(md, d, p, lambda r: num(r["GC_Drill_m"]))
        exc_h = total(fl, d, p, lambda r: num(r["Operating_h"]) if r["Equipment_Class"] == "Excavator" else 0.0)
        trk_h = total(fl, d, p, lambda r: num(r["Operating_h"]) if r["Equipment_Class"] == "Haul Truck" else 0.0)
        sold = total(gd, d, p, lambda r: num(r["Gold_oz"]) if r["Type"] == "Sale" else 0.0)
        put("Ore_Mined_t", p, ore_t)
        put("Ore_Mined_oz", p, ore_oz)
        put("Ore_Grade_gpt", p, div(ore_oz * T, ore_graded_t))
        put("Ore_Ungraded_t", p, ore_t - ore_graded_t)
        put("Waste_t", p, waste_t)
        put("TMM_t", p, ore_t + waste_t)
        put("Strip_Ratio", p, div(waste_t, ore_t))
        put("Rehandle_t", p, rehandle)
        put("Unclassified_t", p, moved - ore_t - waste_t - rehandle)
        put("Milled_t", p, mil)
        put("Feed_oz", p, fz)
        put("Recovered_oz", p, rz)
        put("Head_Grade_gpt", p, div(fz * T, mil_head))
        put("Tails_Grade_gpt", p, div(tz * T, mil_tails))
        put("Recovery_pct", p, div(rz, fz_recon))
        put("Gold_Poured_oz", p, poured)
        put("Throughput_tph", p, div(mil_timed, run_h))
        put("Mill_Calendar_h", p, cal_h)
        put("Mill_Availability_pct", p, div(cal_h - plan_h - unpl_h, cal_h))
        put("Mill_Utilisation_pct", p, div(run_h, cal_h - plan_h - unpl_h))
        put("Cyanide_kgpt", p, div(cn_kg, mil))
        put("Hours_Worked", p, hrs)
        put("Recordables", p, rec)
        put("LTI", p, lti)
        put("Gold_Shipped_oz", p, shipped)
        put("GC_Drill_m", p, gc_m)
        put("Gravity_Share_pct", p, div(gravity, rz))
        put("Exc_Productivity_tph", p, div(ore_t + waste_t + rehandle, exc_h))
        put("Trk_Productivity_tph", p, div(ore_t + waste_t + rehandle, trk_h))
        put("Gold_Sold_oz", p, sold)
    # point value: gold in circuit on the day
    gic = total(pl, d, "Day", lambda r: num(r["GIC_oz"]))
    put("GIC_oz", "Day", gic if gic > 0 else None)
    # days since last LTI: prior-year date from Config or the latest LTI day on or before d
    prior = S.b.sample.get("prior_lti")
    lti_days = [r["Date"] for r in sf if num(r["LTI"]) > 0 and r["Date"] <= d]
    cands = [x for x in (prior, max(lti_days) if lti_days else None) if x]
    put("Days_Since_LTI", "Day", float((d - max(cands)).days) if cands else None)
    # rolling 12 months + month to date: window from the same month one year earlier; history months are used
    # for months in the window before the current month that have no daily hours
    hist = S.rows("tblSafetyHistory")
    ms = date(d.year, d.month, 1)
    win = date(ms.year - 1, ms.month, 1)

    def use_hist(h):
        m = h["Month"]
        return win <= m < ms and not any(hours(r) > 0 for r in sf if (r["Date"].year, r["Date"].month) == (m.year, m.month))

    for key, cur_fn, hist_col in (("TRIFR_12m", recordables, "Recordables"), ("LTIFR_12m", lambda r: num(r["LTI"]), "LTI")):
        n = sum(cur_fn(r) for r in sf if win <= r["Date"] <= d) + sum(num(h[hist_col]) for h in hist if use_hist(h))
        dd = sum(hours(r) for r in sf if win <= r["Date"] <= d) + sum(num(h["Hours_Total"]) for h in hist if use_hist(h))
        put(key, "Day", n / dd * 1_000_000 if dd > 0 else None)
    return out


def expected_budgets(S: Sample, d: date) -> dict[str, object]:
    """Pro-rata budgets from tblBudget: month x day-of-month / days in month; YTD = prior full months + MTD."""
    T = S.troy
    bud = {r["Month"]: r for r in S.rows("tblBudget")}
    dim = calendar.monthrange(d.year, d.month)[1]

    def month_val(m, fn):
        r = bud.get(date(d.year, m, 1))
        return fn(r) if r else 0.0

    def mtd(fn):
        return month_val(d.month, fn) * d.day / dim

    def ytd(fn):
        return sum(month_val(m, fn) for m in range(1, d.month)) + mtd(fn)

    milled = lambda r: num(r["Milled_t"])  # noqa: E731
    ore_t = lambda r: num(r["Ore_Mined_t"])  # noqa: E731
    ore_oz = lambda r: num(r["Ore_Mined_t"]) * num(r["Ore_Grade_gpt"]) / T  # noqa: E731
    feed_oz = lambda r: num(r["Milled_t"]) * num(r["Head_Grade_gpt"]) / T  # noqa: E731
    rec_oz = lambda r: feed_oz(r) * num(r["Recovery_pct"])  # noqa: E731
    return {
        "Bud_Milled_t_MTD": mtd(milled),
        "Bud_Milled_t_YTD": ytd(milled),
        "Bud_Ore_Grade_gpt_MTD": div(mtd(ore_oz) * T, mtd(ore_t)),
        "Bud_Recovery_pct_YTD": div(ytd(rec_oz), ytd(feed_oz)),
    }


def expected_stockpiles(S: Sample, rd: date) -> list[dict]:
    """Balance window per stockpile: from the later of 1 January and its Survey_Date (opening = start of that day) to rd."""
    T = S.troy
    out = []
    for sp in S.rows("tblStockpiles"):
        nm, op_t, op_g = sp["Stockpile"], num(sp["Opening_t"]), num(sp["Opening_gpt"])
        since = max(date(rd.year, 1, 1), as_date(sp["Survey_Date"]) or date(rd.year, 1, 1))
        mv = [r for r in S.rows("tblMovement") if since <= r["Date"] <= rd]
        pl = [r for r in S.rows("tblPlant") if since <= r["Date"] <= rd]
        milled = sum(num(r["Milled_t"]) for r in pl)
        feed_oz = sum(num(r["Milled_t"]) * num(r["Head_Grade_gpt"]) / T for r in pl)
        tip = [r for r in mv if S.dst_type.get(r["Destination"]) == "Plant"]
        tip_t = sum(float(r["Tonnes_t"]) for r in tip)
        tip_oz = sum(S.oz(r) for r in tip)
        feeds = sp["Plant_Feed"] == "Y"
        ins = [r for r in mv if r["Destination"] == nm]
        outs = [r for r in mv if r["Source"] == nm]
        in_t, in_oz = sum(float(r["Tonnes_t"]) for r in ins), sum(S.oz(r) for r in ins)
        out_t, out_oz = sum(float(r["Tonnes_t"]) for r in outs), sum(S.oz(r) for r in outs)
        pf_t = milled - tip_t if feeds else 0.0
        pf_oz = feed_oz - tip_oz if feeds else 0.0
        bal = op_t + in_t - out_t - pf_t
        bal_oz = op_t * op_g / T + in_oz - out_oz - pf_oz
        out.append({"name": nm, "op_t": op_t, "op_g": op_g, "in_t": in_t, "out_t": out_t, "pf_t": pf_t, "bal": bal,
                    "grade": bal_oz * T / bal if bal > 0 else None})
    return out


# ----------------------------------------------------------------------------- workbook readers
class CalcDaily:
    def __init__(self, ws):
        rows = list(ws.iter_rows(values_only=True))
        self.hdr = {h: i for i, h in enumerate(rows[0]) if h}
        self.by_date = {}
        for row in rows[1:]:
            d = as_date(row[0])
            if d:
                self.by_date[d] = row

    def get(self, d: date, hdr: str):
        if hdr not in self.hdr:
            return f"<no column {hdr}>"
        row = self.by_date.get(d)
        return norm(row[self.hdr[hdr]]) if row else f"<no row {d}>"


def config_values(wb) -> dict:
    ws = wb["Config"]
    out = {}
    for r in range(1, 40):
        nm = ws.cell(r, 3).value
        if isinstance(nm, str) and nm.startswith("cfg_"):
            out[nm] = ws.cell(r, 2).value
    return out


def last_data_date(cfg: dict):
    d = as_date(cfg.get("cfg_LastDataDate"))
    return d if d and d.year >= 1901 else None


def find_row(ws, col: int, text: str, start=1):
    for r in range(start, ws.max_row + 1):
        v = ws.cell(r, col).value
        if isinstance(v, str) and v.startswith(text):
            return r
    return None


def scan_errors(wb, allowed) -> dict[str, list[str]]:
    found: dict[str, list[str]] = {}
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if isinstance(v, str) and v in ERROR_VALUES and not allowed(ws.title, c):
                    found.setdefault(ws.title, []).append(f"{c.coordinate}={v}")
    return found


def chart_gap_ok(cd, src: str, d: date, last) -> bool:
    """A guarded Chart_Data cell may be #N/A after the last data date, or where the Calc_Daily value is not a number (blank ratio)."""
    return last is None or d > last or (cd is not None and not is_num(cd.get(d, src)))


def chart_guard_allowed(wb, series, last, cd=None):
    """Only #N/A in guarded Chart_Data columns is acceptable, and only where chart_gap_ok says so."""
    ws = wb["Chart_Data"]
    guarded = {j + 1: s[1] for j, s in enumerate(series) if s[2]}

    def allowed(sheet, c):
        if sheet != "Chart_Data" or c.value != "#N/A" or c.column not in guarded:
            return False
        d = as_date(ws.cell(c.row, 1).value)
        return d is not None and chart_gap_ok(cd, guarded[c.column], d, last)

    return allowed


# ----------------------------------------------------------------------------- checks
def check_calc_daily(C: Checks, S: Sample, cd: CalcDaily, dates=None, tag="Calc_Daily", budgets=True):
    for d in dates or CHECK_DATES:
        exp = expected_kpis(S, d)
        for key, periods in exp.items():
            for period, val in periods.items():
                hdr = key if period == "Day" else f"{key}_{period}"
                C.num(f"{tag} {d} {hdr}", val, cd.get(d, hdr), TOL_SUM if KIND[key] else TOL_RATIO)
        if budgets:
            for hdr, val in expected_budgets(S, d).items():
                C.num(f"{tag} {d} {hdr}", val, cd.get(d, hdr), TOL_RATIO)


def expected_checks_line(S: Sample, rd: date) -> str:
    """The Daily_Report B7 completeness line, from the sample rows."""
    def n(table, cond=lambda r: True):
        return sum(1 for r in S.rows(table) if r["Date"] == rd and cond(r))
    mv = S.rows("tblMovement")
    unknown = sum(1 for r in mv if r["Source"] and r["Source"] not in S.src_type) + sum(1 for r in mv if r["Destination"] and r["Destination"] not in S.dst_type)
    outside = sum(1 for r in mv if r["Date"].year != rd.year)
    return (f"Rows for this date: Plant {n('tblPlant', lambda r: r['Milled_t'] is not None)}, Movements {n('tblMovement')}, Fleet {n('tblFleet')}, "
            f"Mining daily {n('tblMiningDaily', lambda r: r['Diesel_L'] is not None)}, Safety {n('tblSafety', lambda r: r['Hours_Employees'] is not None)}, "
            f"Comments {n('tblCommentary')}  |  Movement rows with unknown source or destination: {unknown}  |  Movement rows dated outside the year: {outside}")


def var(a, b):
    return a / b - 1 if is_num(a) and is_num(b) and b != 0 else None


def check_daily_report(C: Checks, S: Sample, b, wb, cd: CalcDaily, rd: date, tag="Daily_Report"):
    ws = wb["Daily_Report"]
    C.eq(f"{tag} C5 report date", rd, ws["C5"].value)
    C.eq(f"{tag} rpt_Row (L5) = day of year", float((rd - date(YEAR, 1, 1)).days + 1), ws["L5"].value)
    C.eq(f"{tag} B7 completeness line (rpt_Checks)", expected_checks_line(S, rd), ws["B7"].value)
    C.eq(f"{tag} F5 warning blank (date in year, year check OK)", None, norm(ws["F5"].value))
    C.add(f"{tag} B6 header shows data-to date and week", "Data to ... | Week ...", fmt(ws["B6"].value),
          isinstance(ws["B6"].value, str) and ws["B6"].value.startswith("Data to ") and "Week " in ws["B6"].value)
    C.eq(f"{tag} cfg_FeedCheck (exactly one Plant_Feed = Y)", "OK", config_values(wb).get("cfg_FeedCheck"))
    r0 = find_row(ws, 2, "STOCKPILES")
    C.eq(f"{tag} plant feed warning line blank", None, ws.cell(r0 - 1, 2).value if r0 else "<no STOCKPILES block>")
    rows: dict[str, int] = {}
    for r in range(1, ws.max_row + 1):
        v = ws.cell(r, 2).value
        if isinstance(v, str) and v not in rows:
            rows[v] = r
    for section, items in bw.REPORT_LAYOUT:
        mism, n = [], 0
        for key, show_budget in items:
            k = bw.KPI[key]
            r = rows.get(k["label"])
            if r is None:
                mism.append(f"row '{k['label']}' missing")
                continue
            has_p = k["kind"] in ("sum", "ratio")
            has_b = show_budget and b.has_budget(k)
            want = {
                4: cd.get(rd, key),
                5: cd.get(rd, key + "_MTD") if has_p else None,
                8: cd.get(rd, key + "_YTD") if has_p else None,
                6: cd.get(rd, f"Bud_{key}_MTD") if has_b else None,
                9: cd.get(rd, f"Bud_{key}_YTD") if has_b else None,
            }
            want[7] = var(want[5], want[6]) if has_b else None
            want[10] = var(want[8], want[9]) if has_b else None
            for col, w in want.items():
                n += 1
                got = norm(ws.cell(r, col).value)
                if not close(w, got, TOL_SAME):
                    mism.append(f"{k['label']} {get_column_letter(col)}{r} exp {fmt(w)} got {fmt(got)}")
        C.group(f"{tag} {section} cells vs Calc_Daily", n, mism)
    # stockpiles, recomputed independently
    r0 = find_row(ws, 2, "STOCKPILES")
    exp_sp = expected_stockpiles(S, rd)
    mism, n = [], 0
    plaus = []
    for i, e in enumerate(exp_sp):
        r = (r0 or 0) + 1 + i
        got = {c: norm(ws.cell(r, c).value) for c in range(2, 11)} if r0 else {}
        for col, key, tol in ((2, "name", 0), (4, "op_t", TOL_SUM), (5, "op_g", TOL_RATIO), (6, "in_t", TOL_SUM), (7, "out_t", TOL_SUM),
                              (8, "pf_t", TOL_SUM), (9, "bal", TOL_SUM), (10, "grade", TOL_RATIO)):
            n += 1
            if not close(e[key], got.get(col), tol):
                mism.append(f"{e['name']} {get_column_letter(col)}{r} exp {fmt(e[key])} got {fmt(got.get(col))}")
        if not (e["bal"] >= 0 and (e["grade"] is None or 0 <= e["grade"] <= 10)):
            plaus.append(f"{e['name']} balance {fmt(e['bal'])} t at {fmt(e['grade'])} g/t")
    C.group(f"{tag} stockpile block (independent)", n, mism)
    C.group(f"{tag} stockpile balances plausible (t >= 0, 0 to 10 g/t)", len(exp_sp), plaus)
    # commentary in entry order
    r0 = find_row(ws, 2, "COMMENTARY")
    comments = [(c["Area"], c["Comment"]) for c in S.rows("tblCommentary") if c["Date"] == rd]
    mism, n = [], 0
    for i in range(bw.COMMENT_LINES):
        e = comments[i] if i < len(comments) else (None, None)
        r = (r0 or 0) + 1 + i
        got = (norm(ws.cell(r, 2).value), norm(ws.cell(r, 3).value)) if r0 else (None, None)
        n += 1
        if got != e:
            mism.append(f"line {i + 1} exp {e} got {got}")
    C.add(f"{tag} commentary lines for report date", f"{len(comments)} comments, entry order", f"{len(comments)} comments" if not mism else mism[0], not mism)
    C.add(f"{tag} commentary count fits the block", f"<= {bw.COMMENT_LINES}", str(len(comments)), len(comments) <= bw.COMMENT_LINES)


def check_monthly(C: Checks, S: Sample, b, wb, cd: CalcDaily, last: date):
    ws = wb["Monthly_Summary"]
    actual_rows: dict[str, int] = {}
    for r in range(1, ws.max_row + 1):
        v, kind = ws.cell(r, 2).value, ws.cell(r, 3).value
        if isinstance(v, str) and kind == "Actual" and v not in actual_rows:
            actual_rows[v] = r
    # Milled_t, independent of the workbook
    r = actual_rows[bw.KPI["Milled_t"]["label"]]
    pl = S.rows("tblPlant")
    bud = {row["Month"]: row for row in S.rows("tblBudget")}
    # budget: full month, except the month in progress (pro-rata to the last data date); year = YTD budget at that date
    exp_act, exp_bud = {}, {}
    for m in range(1, 13):
        ms = date(YEAR, m, 1)
        exp_act[m] = sum(num(p["Milled_t"]) for p in pl if p["Date"].month == m and p["Date"] <= last) if ms <= last else None
        full = num(bud[ms]["Milled_t"]) if ms in bud else 0.0
        exp_bud[m] = full * last.day / calendar.monthrange(YEAR, m)[1] if m == last.month else full
    for m in (1, 2, 3, 4):
        C.num(f"Monthly_Summary Milled {MONTHS[m - 1]} actual", exp_act[m], ws.cell(r, 3 + m).value)
    ytd = sum(num(p["Milled_t"]) for p in pl if p["Date"] <= last)
    C.num("Monthly_Summary Milled Year actual = YTD at last data date", ytd, ws.cell(r, 16).value)
    mism_b, mism_v = [], []
    for m in range(1, 13):
        got_b = norm(ws.cell(r + 1, 3 + m).value)
        if not close(exp_bud[m], got_b, TOL_RATIO):
            mism_b.append(f"{MONTHS[m - 1]} exp {fmt(exp_bud[m])} got {fmt(got_b)}")
        got_v = norm(ws.cell(r + 2, 3 + m).value)
        ev = var(exp_act[m], exp_bud[m])
        if not close(ev, got_v, TOL_RATIO):
            mism_v.append(f"{MONTHS[m - 1]} exp {fmt(ev)} got {fmt(got_v)}")
    ytd_bud = sum(v for m, v in exp_bud.items() if m <= last.month)
    if not close(ytd_bud, ws.cell(r + 1, 16).value, TOL_RATIO):
        mism_b.append(f"Year exp {fmt(ytd_bud)} got {fmt(ws.cell(r + 1, 16).value)}")
    if not close(var(ytd, ytd_bud), ws.cell(r + 2, 16).value, TOL_RATIO):
        mism_v.append(f"Year exp {fmt(var(ytd, ytd_bud))} got {fmt(ws.cell(r + 2, 16).value)}")
    C.group("Monthly_Summary Milled budget row = tblBudget, pro-rata for the current month (12 months + year)", 13, mism_b)
    C.group("Monthly_Summary Milled variance row = actual / budget - 1", 13, mism_v)
    # every monthly KPI row against Calc_Daily
    mism, n = [], 0
    for key in bw.MONTHLY_KPIS:
        k = bw.KPI[key]
        r = actual_rows.get(k["label"])
        if r is None:
            mism.append(f"row '{k['label']}' missing")
            continue
        per = k["kind"] in ("sum", "ratio")
        for m in range(1, 13):
            ms = date(YEAR, m, 1)
            me = min(date(YEAR, m, calendar.monthrange(YEAR, m)[1]), last)
            want = cd.get(me, f"{key}_MTD" if per else key) if ms <= last else None
            n += 1
            got = norm(ws.cell(r, 3 + m).value)
            if not close(want, got, TOL_SAME):
                mism.append(f"{k['label']} {MONTHS[m - 1]} exp {fmt(want)} got {fmt(got)}")
            if b.has_budget(k):
                n += 1
                wb_ = cd.get(me, f"Bud_{key}_MTD") if ms <= last else cd.get(ms, f"Bud_{key}_Month")
                got_b = norm(ws.cell(r + 1, 3 + m).value)
                if not close(wb_, got_b, TOL_SAME):
                    mism.append(f"{k['label']} {MONTHS[m - 1]} budget exp {fmt(wb_)} got {fmt(got_b)}")
        n += 1
        want = cd.get(last, f"{key}_YTD" if per else key)
        got = norm(ws.cell(r, 16).value)
        if not close(want, got, TOL_SAME):
            mism.append(f"{k['label']} Year exp {fmt(want)} got {fmt(got)}")
    C.group("Monthly_Summary all KPI rows vs Calc_Daily", n, mism)


def check_chart_data(C: Checks, wb, series, last: date, cd=None):
    ws = wb["Chart_Data"]
    rows = list(ws.iter_rows(values_only=True))
    for j, (name, src, guard) in enumerate(series):
        if j == 0:
            continue
        bad = []
        for row in rows[1:]:
            d = as_date(row[0])
            if d is None:
                continue
            v = row[j]
            if guard and (d > last or (cd is not None and not is_num(cd.get(d, src)))):
                if v != "#N/A":
                    bad.append(f"{d} {fmt(v)}")
            elif not is_num(v):
                bad.append(f"{d} {fmt(v)}")
        label = "NA after last data date or blank KPI, numeric before" if guard else "numeric all year"
        C.group(f"Chart_Data {name} ({label})", len(rows) - 1, bad)


def check_inspect(C: Checks, built: Path, outdir: Path):
    cmd = [sys.executable, str(REPO / "tools" / "xl_inspect.py"), str(built), "-o", str(outdir)]
    res = subprocess.run(cmd, capture_output=True, text=True, timeout=600)
    js = outdir / (built.stem + ".json")
    if res.returncode != 0 or not js.exists():
        C.add("xl_inspect ran", "exit 0 and JSON written", f"exit {res.returncode}: {res.stderr.strip()[-100:]}", False)
        return
    info = json.loads(js.read_text(encoding="utf-8"))
    C.add("xl_inspect external links", "0", str(len(info.get("external_links", []))), not info.get("external_links"))
    C.add("xl_inspect VBA project", "none", "present" if info.get("has_vba_project") else "none", not info.get("has_vba_project"))
    dang = info.get("dangling_sheet_refs", [])
    C.add("xl_inspect dangling sheet references", "0", str(len(dang)) + (f" ({', '.join(dang[:3])})" if dang else ""), not dang)
    broken = [d["name"] for d in info.get("defined_names", []) if d.get("broken")]
    C.add("xl_inspect broken defined names", "0", str(len(broken)) + (f" ({', '.join(broken[:3])})" if broken else ""), not broken)
    C.add("xl_inspect scan errors", "none", "; ".join(info.get("errors", [])) or "none", not info.get("errors"))


def check_blank(C: Checks, wb, series):
    cfg = config_values(wb)
    last = last_data_date(cfg)
    C.add("Blank cfg_LastDataDate", "no data (blank or zero)", fmt(cfg.get("cfg_LastDataDate")), last is None)
    found = scan_errors(wb, chart_guard_allowed(wb, series, last))
    C.add("Blank workbook error values outside the Chart_Data guard", "0",
          str(sum(len(v) for v in found.values())) + ("; ".join(f"{k}: {v[0]}" for k, v in found.items())[:60] if found else ""), not found)
    ws = wb["Daily_Report"]
    title = norm(ws["B2"].value)
    C.add("Blank Daily_Report title renders", "text", fmt(title), isinstance(title, str) and bool(title))
    C.eq("Blank cfg_FeedCheck (schema stockpiles)", "OK", cfg.get("cfg_FeedCheck"))
    C.add("Blank Daily_Report B7 completeness line all zero", "Plant 0 ... outside the year: 0", fmt(ws["B7"].value),
          isinstance(ws["B7"].value, str) and ws["B7"].value.startswith("Rows for this date: Plant 0, Movements 0, Fleet 0, Mining daily 0, Safety 0, Comments 0")
          and ws["B7"].value.endswith("outside the year: 0"))
    today = date.today()
    exp_default = max(date(YEAR, 1, 1), min(today - timedelta(days=1), date(YEAR, 12, 31)))
    c5 = as_date(ws["C5"].value)
    C.eq("Blank Daily_Report default date = clamp(today - 1) in year", exp_default, c5)
    C.eq("Blank Daily_Report rpt_Row resolves", float((exp_default - date(YEAR, 1, 1)).days + 1), ws["L5"].value)
    bad, n = [], 0
    for section, items in bw.REPORT_LAYOUT:
        for key, _sb in items:
            r = find_row(ws, 2, bw.KPI[key]["label"])
            if r is None:
                bad.append(f"row '{bw.KPI[key]['label']}' missing")
                continue
            for col in range(4, 11):
                n += 1
                v = norm(ws.cell(r, col).value)
                if v is not None and v != 0:
                    bad.append(f"{get_column_letter(col)}{r}={fmt(v)}")
    C.group("Blank Daily_Report KPI cells blank or zero", n, bad)


# ----------------------------------------------------------------------------- edge workbook and list probe
def list_catalogue(schema: dict) -> dict[str, list]:
    """Every lst_ name the generator defines: schema lists, derived lists and the key column of Lists-sheet reference tables."""
    out = dict(schema["lists"])
    out.update(schema.get("derived_lists", {}))
    for rdef in schema["reference_tables"].values():
        if rdef["sheet"] == "Lists":
            out[rdef["key"]] = [row[0] for row in rdef["rows"]]
    return out


def make_edge(b, out: Path) -> Sample:
    """Copy the demo builder's workbook and sample with partially entered rows, add the lst_ probe sheet, save."""
    import copy
    smp = copy.deepcopy(b.sample)
    S = Sample(b, smp)
    wb = b.wb
    edits = {  # table -> (row date, {field: value})
        "tblPlant": [(date(YEAR, 2, 5), {"Tails_Grade_gpt": None}), (date(YEAR, 2, 6), {"Head_Grade_gpt": None}),
                     (date(YEAR, 2, 7), {"Mill_Run_h": None}),
                     (date(YEAR, 2, 8), {"Milled_t": 0, "Mill_Run_h": None, "Mill_Planned_Maint_h": 24, "Mill_Unplanned_Down_h": 0})],
    }
    for tbl, items in edits.items():
        info = b.tables[tbl]
        ws = wb[info["sheet"]]
        names = S.fields[tbl]
        for d, vals in items:
            r = info["first"] + (d - date(YEAR, 1, 1)).days
            row = smp[tbl][(d - date(YEAR, 1, 1)).days]
            assert row[0] == d
            for k, v in vals.items():
                row[names.index(k)] = v
                ws[f"{info['cols'][k]}{r}"] = v
    # ungraded ore row on 9 Feb (first Petowal Pit Ore HG row of that day), an unknown source and a prior-year row
    info = b.tables["tblMovement"]
    ws = wb[info["sheet"]]
    names = S.fields["tblMovement"]
    mv = smp["tblMovement"]
    i9 = next(i for i, row in enumerate(mv) if row[0] == date(YEAR, 2, 9) and row[3] == "Ore HG")
    mv[i9][names.index("Grade_gpt")] = None
    ws[f"{info['cols']['Grade_gpt']}{info['first'] + i9}"] = None
    extra = [[date(YEAR - 1, 12, 31), None, "Petowal Pit", "Ore HG", "ROM Pad", 777, 2.0, None, None, None],
             [date(YEAR, 2, 10), None, "Unknown Pit", "Ore HG", "ROM Pad", 1000, 2.0, None, None, None]]
    for row in extra:
        r = info["first"] + len(mv)
        mv.append(row)
        for k, v in zip(names, row):
            if v is not None:
                ws[f"{info['cols'][k]}{r}"] = v
    # re-surveyed stockpile: Stockpile LG opening reset on 1 Feb (movements and plant rows before that day leave its balance)
    info = b.tables["tblStockpiles"]
    ws = wb[info["sheet"]]
    i_lg = next(i for i, row in enumerate(smp["tblStockpiles"]) if row[0] == "Stockpile LG")
    names = S.fields["tblStockpiles"]
    for k, v in (("Opening_t", 160000), ("Opening_gpt", 0.58), ("Survey_Date", date(YEAR, 2, 1))):
        smp["tblStockpiles"][i_lg][names.index(k)] = v
        ws[f"{info['cols'][k]}{info['first'] + i_lg}"] = v
    # probe sheet: size, first and last entry of every drop-down name
    pr = wb.create_sheet("Probe")
    pr.append(["name", "rows", "first", "last"])
    for lname in list_catalogue(b.s):
        pr.append([lname, f"=ROWS(lst_{lname})", f"=INDEX(lst_{lname},1)", f"=INDEX(lst_{lname},ROWS(lst_{lname}))"])
    wb.save(out)
    S._cache.clear()
    return S


def check_lists(C: Checks, b, wb_edge):
    """Every lst_ name covers exactly the schema list (F001), and every list field validates against its name."""
    cat = list_catalogue(b.s)
    got = {row[0]: row[1:] for row in wb_edge["Probe"].iter_rows(min_row=2, values_only=True) if row[0]}
    bad = []
    for lname, values in cat.items():
        g = got.get(lname)
        exp = (float(len(values)), values[0], values[-1])
        if g is None or tuple(norm(x) for x in g) != exp:
            bad.append(f"lst_{lname} exp {exp} got {g}")
    C.group("Lists: ROWS/first/last of every lst_ name = schema list (LibreOffice)", len(cat), bad)
    bad, n = [], 0
    fields = [(t["table"], f) for t in b.s["tables"] for f in t["fields"]]
    fields += [(rdef["table"], c) for rdef in b.s["reference_tables"].values() for c in rdef["columns"]]
    for tbl, f in fields:
        if f.get("type") != "list":
            continue
        n += 1
        info = b.tables[tbl]
        ws = b.wb[info["sheet"]]
        want = f"=lst_{f['list']}"
        col_ref = f"{info['cols'][f['name']]}{info['first']}"
        dvs = [dv for dv in ws.data_validations.dataValidation if dv.formula1 == want and col_ref in dv.sqref]
        if not dvs:
            bad.append(f"{tbl}[{f['name']}] has no list validation {want} on {col_ref}")
        elif f"lst_{f['list']}" not in b.wb.defined_names:
            bad.append(f"{tbl}[{f['name']}]: name lst_{f['list']} missing")
    C.group("Lists: every list field validates against an existing lst_ name", n, bad)
    req = [f"{r}.Type" for r, rdef in b.s["reference_tables"].items() if r in ("Sources", "Destinations")
           and not any(c["name"] == "Type" and c.get("required") for c in rdef["columns"])]
    C.group("Schema: Sources.Type and Destinations.Type are required", 2, [f"{x} not required" for x in req])


# ----------------------------------------------------------------------------- driver
def recalc(files: list[Path], outdir: Path, soffice: str) -> list[Path]:
    outdir.mkdir(parents=True, exist_ok=True)
    profile = (outdir.parent / "lo_profile").resolve()
    cmd = [soffice, f"-env:UserInstallation=file://{profile}", "--headless", "--calc", "--convert-to", "xlsx", "--outdir", str(outdir)] + [str(f) for f in files]
    res = subprocess.run(cmd, capture_output=True, text=True, timeout=900)
    out = [outdir / f.name for f in files]
    missing = [p for p in out if not p.exists()]
    if missing:
        sys.exit(f"LibreOffice did not produce {', '.join(str(m) for m in missing)}\n{res.stdout}\n{res.stderr}")
    return out


def load(path: Path):
    return openpyxl.load_workbook(path, data_only=True)


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--keep", action="store_true", help="keep the built and recalculated files")
    ap.add_argument("--workdir", default=None, help="folder for the built files (default: a fresh temp folder)")
    ap.add_argument("--soffice", default=shutil.which("soffice") or shutil.which("libreoffice") or "soffice", help="LibreOffice executable")
    ap.add_argument("--schema", default=str(bw.DEFAULT_SCHEMA))
    ap.add_argument("--year", type=int, default=YEAR, help="reporting year to build and check (2028 exercises the leap-year branch)")
    a = ap.parse_args(argv)
    set_year(a.year)

    t0 = time.time()
    auto = a.workdir is None
    work = Path(a.workdir) if a.workdir else Path(tempfile.mkdtemp(prefix="mps_validate_"))
    work.mkdir(parents=True, exist_ok=True)
    built = work / "built"
    built.mkdir(exist_ok=True)
    schema = Path(a.schema)
    C = Checks()
    try:
        print(f"work folder: {work}")
        print("building workbooks ...", flush=True)
        f_demo, f_default, f_blank, f_edge = built / "MPS_demo.xlsx", built / "MPS_demo_default_date.xlsx", built / "MPS_blank.xlsx", built / "MPS_edge.xlsx"
        b = bw.build(schema, f_demo, YEAR, SAMPLE_DAYS, REPORT_DATE)
        bw.build(schema, f_default, YEAR, SAMPLE_DAYS, None)
        bw.build(schema, f_blank, YEAR, 0, None)
        S = Sample(b)
        S_edge = make_edge(b, f_edge)   # after Sample(b): the edge edits the builder's workbook in place
        last = S.last_plant_date()
        C.eq("Sample last day with data", CHECK_DATES[-1], last)
        # the Config cell must hold the formula, not a pasted date (the earlier override bug)
        ref = b.wb.defined_names["cfg_LastDataDate"].attr_text.split("!")[1].replace("$", "")
        cell = b.wb["Config"][ref].value
        C.add("Config cfg_LastDataDate is a formula", "=...", str(cell)[:30], isinstance(cell, str) and cell.startswith("="))
        dvs = [dv for dv in b.wb["Daily_Report"].data_validations.dataValidation if "C5" in str(dv.sqref)]
        C.add("Daily_Report C5 stop validation: date within the year", "date, stop, cfg_YearStart..cfg_YearEnd",
              f"{dvs[0].type}, {dvs[0].errorStyle}, {dvs[0].formula1}..{dvs[0].formula2}" if dvs else "none",
              bool(dvs) and dvs[0].type == "date" and dvs[0].errorStyle == "stop" and dvs[0].formula1 == "=cfg_YearStart" and dvs[0].formula2 == "=cfg_YearEnd")
        C.add("Daily_Report C5 default formula matches the restore hint", "same text", "same" if dvs and bw.DEFAULT_DATE_FORMULA in (dvs[0].prompt or "") else "differs",
              bool(dvs) and bw.DEFAULT_DATE_FORMULA in (dvs[0].prompt or ""))
        for t in S.fields:
            S.rows(t)
        C.group("Sample rows match schema field order (width)", len(S.fields), S.width_problems)

        print("recalculating with LibreOffice ...", flush=True)
        r_demo, r_default, r_blank, r_edge = recalc([f_demo, f_default, f_blank, f_edge], work / "recalc", a.soffice)

        print("checking the DEMO workbook ...", flush=True)
        wb = load(r_demo)
        cfg = config_values(wb)
        C.eq("Recalculated cfg_LastDataDate = last plant date", last, last_data_date(cfg))
        cd = CalcDaily(wb["Calc_Daily"])
        check_calc_daily(C, S, cd)
        check_daily_report(C, S, b, wb, cd, REPORT_DATE)
        check_monthly(C, S, b, wb, cd, last)
        check_chart_data(C, wb, b.chart_series, last, cd)
        found = scan_errors(wb, chart_guard_allowed(wb, b.chart_series, last, cd))
        C.add("DEMO error values outside the Chart_Data guard", "0",
              str(sum(len(v) for v in found.values())) + (" " + "; ".join(f"{k}: {v[0]}" for k, v in found.items())[:60] if found else ""), not found)
        wb.close()

        print("checking the default report date ...", flush=True)
        wb = load(r_default)
        today = date.today()
        exp_default = max(date(YEAR, 1, 1), min(today - timedelta(days=1), date(YEAR, 12, 31), last))
        C.eq("Default report date = min(today - 1, last data date) in year", exp_default, as_date(wb["Daily_Report"]["C5"].value))
        wb.close()

        print("checking the blank workbook ...", flush=True)
        wb = load(r_blank)
        check_blank(C, wb, b.chart_series)
        wb.close()

        print("checking the edge workbook (partial rows) and the drop-down lists ...", flush=True)
        wb = load(r_edge)
        cd = CalcDaily(wb["Calc_Daily"])
        check_calc_daily(C, S_edge, cd, EDGE_DATES, "Edge Calc_Daily", budgets=False)
        check_daily_report(C, S_edge, b, wb, cd, REPORT_DATE, tag="Edge Daily_Report")
        found = scan_errors(wb, chart_guard_allowed(wb, b.chart_series, last, cd))
        C.add("Edge error values outside the Chart_Data guard", "0",
              str(sum(len(v) for v in found.values())) + (" " + "; ".join(f"{k}: {v[0]}" for k, v in found.items())[:60] if found else ""), not found)
        check_lists(C, b, wb)
        wb.close()

        print("running xl_inspect ...", flush=True)
        check_inspect(C, f_demo, work / "inspect")
    finally:
        print()
        C.print()
        print(f"elapsed {time.time() - t0:.0f} s")
        if auto and not a.keep:
            shutil.rmtree(work, ignore_errors=True)
        else:
            print(f"files kept in {work}")
    return 1 if C.failed else 0


if __name__ == "__main__":
    sys.exit(main())
'@
Write-TextFile (Join-Path $Root "01_Tools\mps\validate_workbook.py") $content_mps_validate_workbook_py $false
$content_mps_schema_mps_schema_json = @'
{
  "system": "Mako Production System",
  "version": "0.3.0",
  "site": {
    "name": "Mako Gold Mine",
    "company": "Petowal Mining Company S.A.",
    "owner": "Resolute Mining Limited",
    "country": "Senegal",
    "timezone": "GMT"
  },
  "constants": {
    "TROY_OZ_G": 31.1035,
    "HOURS_PER_DAY": 24,
    "RATE_BASIS_HOURS": 1000000
  },
  "lists": {
    "Shift": ["Day", "Night"],
    "Material": ["Ore HG", "Ore MG", "Ore LG", "Mineralised Waste", "Waste", "Topsoil"],
    "Area": ["Safety", "Mining", "Geology", "Processing", "Maintenance", "Environment", "Community", "Supply", "People", "Projects", "General"],
    "Equipment_Class": ["Excavator", "Haul Truck", "Drill", "Dozer", "Grader", "Water Cart", "Wheel Loader", "Ancillary"],
    "Gold_Movement_Type": ["Shipment", "Sale"],
    "Yes_No": ["Y", "N"]
  },
  "reference_tables": {
    "Sources": {
      "table": "tblSources",
      "sheet": "Lists",
      "key": "Source",
      "columns": [
        {"name": "Source", "type": "text", "desc": "Origin of a material movement: a pit or a stockpile"},
        {"name": "Type", "type": "list", "list": "Source_Type", "required": true, "required_when": "{Source}<>\"\"", "desc": "Pit = ex-pit (counts to TMM and strip ratio); Stockpile = rehandle. Required: a source without a type is excluded from every KPI"}
      ],
      "rows": [
        ["Petowal Pit", "Pit"],
        ["Satellite Pit", "Pit"],
        ["ROM Pad", "Stockpile"],
        ["Stockpile HG", "Stockpile"],
        ["Stockpile MG", "Stockpile"],
        ["Stockpile LG", "Stockpile"],
        ["Mineralised Waste Dump", "Stockpile"]
      ]
    },
    "Destinations": {
      "table": "tblDestinations",
      "sheet": "Lists",
      "key": "Destination",
      "columns": [
        {"name": "Destination", "type": "text", "desc": "Where the material was delivered"},
        {"name": "Type", "type": "list", "list": "Destination_Type", "required": true, "required_when": "{Destination}<>\"\"", "desc": "Stockpile = enters a stockpile balance; Plant = direct tip to crusher; Waste = dump. Required"}
      ],
      "rows": [
        ["ROM Pad", "Stockpile"],
        ["Crusher Direct Tip", "Plant"],
        ["Stockpile HG", "Stockpile"],
        ["Stockpile MG", "Stockpile"],
        ["Stockpile LG", "Stockpile"],
        ["Mineralised Waste Dump", "Stockpile"],
        ["Waste Dump", "Waste"],
        ["TSF Wall", "Waste"],
        ["Topsoil Stockpile", "Waste"]
      ]
    },
    "Stockpiles": {
      "table": "tblStockpiles",
      "sheet": "Config",
      "key": "Stockpile",
      "spare_rows": 5,
      "columns": [
        {"name": "Stockpile", "type": "text", "desc": "Stockpile name; must match a Source and a Destination"},
        {"name": "Opening_t", "type": "number", "unit": "t", "min": 0, "max": 5000000, "desc": "Surveyed tonnes at the start of Survey_Date (1 January if blank); movements before that day are not part of the balance"},
        {"name": "Opening_gpt", "type": "number", "unit": "g/t", "min": 0, "max": 50, "desc": "Grade of the opening balance"},
        {"name": "Plant_Feed", "type": "list", "list": "Yes_No", "desc": "Y for exactly one stockpile: the plant draws its feed from it (milled tonnes less direct tip are deducted, so its balance includes ore tipped to the crusher but not yet milled)"},
        {"name": "Survey_Date", "type": "date", "desc": "Day the opening balance applies to (balance at the start of that day; blank = 1 January). After a re-survey enter the new tonnes, grade and date here"}
      ],
      "rows": [
        ["ROM Pad", 0, 0, "Y", null],
        ["Stockpile HG", 0, 0, "N", null],
        ["Stockpile MG", 0, 0, "N", null],
        ["Stockpile LG", 0, 0, "N", null],
        ["Mineralised Waste Dump", 0, 0, "N", null]
      ]
    }
  },
  "derived_lists": {
    "Source_Type": ["Pit", "Stockpile"],
    "Destination_Type": ["Stockpile", "Plant", "Waste"]
  },
  "tables": [
    {
      "table": "tblPlant",
      "sheet": "Plant",
      "title": "Processing plant, one row per day",
      "owner": "Processing (met accountant / shift supervisor)",
      "prefill": "dates",
      "fields": [
        {"name": "Date", "label": "Date", "type": "date", "key": true, "desc": "Production day (day shift start to night shift end)"},
        {"name": "Crushed_t", "label": "Crushed", "unit": "t", "type": "number", "min": 0, "max": 50000, "desc": "Dry tonnes through the primary crusher"},
        {"name": "Milled_t", "label": "Milled", "unit": "t", "type": "number", "min": 0, "max": 50000, "desc": "Dry tonnes milled (mill feed weightometer)"},
        {"name": "Head_Grade_gpt", "label": "Head grade", "unit": "g/t", "type": "number", "min": 0, "max": 50, "desc": "Mill feed grade from the daily metallurgical balance"},
        {"name": "Tails_Grade_gpt", "label": "Tails grade", "unit": "g/t", "type": "number", "min": 0, "max": 10, "desc": "Final tails grade"},
        {"name": "Gravity_Gold_oz", "label": "Gravity gold", "unit": "oz", "type": "number", "min": 0, "max": 5000, "desc": "Gold recovered by the gravity circuit (part of total recovered)"},
        {"name": "Gold_Poured_oz", "label": "Gold poured", "unit": "oz", "type": "number", "min": 0, "max": 20000, "desc": "Fine gold in dore poured on the day"},
        {"name": "Mill_Run_h", "label": "Mill run hours", "unit": "h", "type": "number", "min": 0, "max": 24, "desc": "Hours the mill was running"},
        {"name": "Mill_Planned_Maint_h", "label": "Planned maintenance", "unit": "h", "type": "number", "min": 0, "max": 24, "desc": "Planned maintenance downtime"},
        {"name": "Mill_Unplanned_Down_h", "label": "Unplanned downtime", "unit": "h", "type": "number", "min": 0, "max": 24, "desc": "Unplanned downtime (breakdowns, power, no feed)"},
        {"name": "Crusher_Run_h", "label": "Crusher run hours", "unit": "h", "type": "number", "min": 0, "max": 24, "desc": "Primary crusher running hours"},
        {"name": "Cyanide_kg", "label": "Cyanide", "unit": "kg", "type": "number", "min": 0, "max": 50000, "desc": "Sodium cyanide consumed"},
        {"name": "Lime_kg", "label": "Lime", "unit": "kg", "type": "number", "min": 0, "max": 100000, "desc": "Lime consumed"},
        {"name": "Grinding_Media_kg", "label": "Grinding media", "unit": "kg", "type": "number", "min": 0, "max": 50000, "desc": "Grinding media added"},
        {"name": "Power_kWh", "label": "Power", "unit": "kWh", "type": "number", "min": 0, "max": 1000000, "desc": "Plant power consumed"},
        {"name": "Raw_Water_m3", "label": "Raw water", "unit": "m3", "type": "number", "min": 0, "max": 100000, "desc": "Raw water drawn"},
        {"name": "GIC_oz", "label": "Gold in circuit", "unit": "oz", "type": "number", "min": 0, "max": 50000, "desc": "Gold in circuit at end of day (point value, not additive)"},
        {"name": "Comments", "label": "Comments", "type": "text", "desc": "Short note; use the Commentary sheet for the report narrative"},
        {"name": "Feed_oz", "label": "Contained gold in feed", "unit": "oz", "type": "number", "calc": "IF(OR({Milled_t}=\"\",{Head_Grade_gpt}=\"\"),\"\",{Milled_t}*{Head_Grade_gpt}/{TROY})", "sql": "milled_t * head_grade_gpt / {TROY}", "desc": "Milled t x head grade / 31.1035; blank until the head grade is entered"},
        {"name": "Tails_oz", "label": "Gold lost to tails", "unit": "oz", "type": "number", "calc": "IF(OR({Milled_t}=\"\",{Tails_Grade_gpt}=\"\"),\"\",{Milled_t}*{Tails_Grade_gpt}/{TROY})", "sql": "milled_t * tails_grade_gpt / {TROY}", "desc": "Milled t x tails grade / 31.1035; blank until the tails grade is entered"},
        {"name": "Recovered_oz", "label": "Gold recovered", "unit": "oz", "type": "number", "calc": "IF(OR({Milled_t}=\"\",{Head_Grade_gpt}=\"\",{Tails_Grade_gpt}=\"\"),\"\",{Feed_oz}-{Tails_oz})", "sql": "milled_t * (head_grade_gpt - tails_grade_gpt) / {TROY}", "desc": "Feed oz minus tails oz; blank until both grades are entered"},
        {"name": "Recovery_pct", "label": "Recovery", "unit": "%", "type": "pct", "calc": "IF(AND(ISNUMBER({Recovered_oz}),{Feed_oz}>0),{Recovered_oz}/{Feed_oz},\"\")", "sql": "CASE WHEN head_grade_gpt > 0 THEN 1 - tails_grade_gpt / head_grade_gpt END", "desc": "1 minus tails grade over head grade; blank until both grades are entered"},
        {"name": "Throughput_tph", "label": "Throughput", "unit": "t/h", "type": "number", "calc": "IF(AND({Milled_t}<>\"\",{Mill_Run_h}>0),{Milled_t}/{Mill_Run_h},\"\")", "sql": "CASE WHEN mill_run_h > 0 THEN milled_t / mill_run_h END", "desc": "Milled t per running hour"},
        {"name": "Availability_pct", "label": "Mill availability", "unit": "%", "type": "pct", "calc": "IF(AND({Mill_Run_h}=\"\",{Mill_Planned_Maint_h}=\"\",{Mill_Unplanned_Down_h}=\"\"),\"\",({HOURS_PER_DAY}-{Mill_Planned_Maint_h}-{Mill_Unplanned_Down_h})/{HOURS_PER_DAY})", "sql": "CASE WHEN COALESCE(mill_run_h, mill_planned_maint_h, mill_unplanned_down_h) IS NOT NULL THEN (24 - COALESCE(mill_planned_maint_h, 0) - COALESCE(mill_unplanned_down_h, 0)) / 24.0 END", "desc": "(24 h minus planned and unplanned downtime) / 24 h; blank only when no mill hours are entered"},
        {"name": "Utilisation_pct", "label": "Mill utilisation", "unit": "%", "type": "pct", "calc": "IF(AND(NOT(AND({Mill_Run_h}=\"\",{Mill_Planned_Maint_h}=\"\",{Mill_Unplanned_Down_h}=\"\")),({HOURS_PER_DAY}-{Mill_Planned_Maint_h}-{Mill_Unplanned_Down_h})>0),{Mill_Run_h}/({HOURS_PER_DAY}-{Mill_Planned_Maint_h}-{Mill_Unplanned_Down_h}),\"\")", "sql": "CASE WHEN COALESCE(mill_run_h, mill_planned_maint_h, mill_unplanned_down_h) IS NOT NULL AND (24 - COALESCE(mill_planned_maint_h, 0) - COALESCE(mill_unplanned_down_h, 0)) > 0 THEN COALESCE(mill_run_h, 0) / (24 - COALESCE(mill_planned_maint_h, 0) - COALESCE(mill_unplanned_down_h, 0)) END", "desc": "Run hours over available hours; blank only when no mill hours are entered"}
      ]
    },
    {
      "table": "tblMovement",
      "sheet": "Mining_Movements",
      "title": "Material movements, one row per source, material and destination per day (or per shift)",
      "owner": "Mining (dispatch / production engineer)",
      "rows": 5000,
      "unique": [["Date", "Shift", "Source", "Material", "Destination"]],
      "fields": [
        {"name": "Date", "label": "Date", "type": "date", "required": true, "desc": "Production day"},
        {"name": "Shift", "label": "Shift", "type": "list", "list": "Shift", "desc": "Optional; leave blank for a daily total"},
        {"name": "Source", "label": "Source", "type": "list", "list": "Source", "required": true, "desc": "Pit or stockpile the material came from (Lists sheet)"},
        {"name": "Material", "label": "Material", "type": "list", "list": "Material", "required": true, "desc": "Ore grade class or waste type"},
        {"name": "Destination", "label": "Destination", "type": "list", "list": "Destination", "required": true, "desc": "Where the material went (Lists sheet)"},
        {"name": "Tonnes_t", "label": "Tonnes", "unit": "t", "type": "number", "min": 0, "max": 200000, "required": true, "desc": "Dry tonnes moved (truck factor or weightometer)"},
        {"name": "Grade_gpt", "label": "Grade", "unit": "g/t", "type": "number", "min": 0, "max": 100, "required_when": "{Is_Ore}=1", "required_when_sql": "material LIKE 'Ore%'", "desc": "Grade control grade; required for ore (an ore row without a grade is highlighted and reported as ungraded tonnes), recommended for mineralised waste"},
        {"name": "BCM", "label": "Volume", "unit": "bcm", "type": "number", "min": 0, "max": 100000, "desc": "Optional bank cubic metres"},
        {"name": "Loads", "label": "Truck loads", "type": "int", "min": 0, "max": 5000, "desc": "Optional truck count"},
        {"name": "Comments", "label": "Comments", "type": "text"},
        {"name": "Source_Type", "label": "Source type", "type": "text", "calc": "IF({Source}=\"\",\"\",IFERROR(INDEX(tblSources[Type],MATCH({Source},tblSources[Source],0)),\"?\"))", "values_from": "Source_Type", "warn_when": "{Source_Type}=\"?\"", "desc": "Pit or Stockpile, looked up from Lists; ? when the source is not in tblSources (row is highlighted and excluded from every KPI)"},
        {"name": "Dest_Type", "label": "Destination type", "type": "text", "calc": "IF({Destination}=\"\",\"\",IFERROR(INDEX(tblDestinations[Type],MATCH({Destination},tblDestinations[Destination],0)),\"?\"))", "values_from": "Destination_Type", "warn_when": "{Dest_Type}=\"?\"", "desc": "Stockpile, Plant or Waste, looked up from Lists; ? when the destination is not in tblDestinations (row is highlighted)"},
        {"name": "Is_Ore", "label": "Ore flag", "type": "int", "calc": "IF({Material}=\"\",\"\",IF(LEFT({Material},3)=\"Ore\",1,0))", "sql": "CASE WHEN material LIKE 'Ore%' THEN 1 ELSE 0 END", "desc": "1 when Material starts with Ore"},
        {"name": "Contained_oz", "label": "Contained gold", "unit": "oz", "type": "number", "calc": "IF(OR({Tonnes_t}=\"\",{Grade_gpt}=\"\"),0,{Tonnes_t}*{Grade_gpt}/{TROY})", "sql": "tonnes_t * COALESCE(grade_gpt, 0) / {TROY}", "desc": "Tonnes x grade / 31.1035, any material with a grade"},
        {"name": "Ore_oz", "label": "Ore gold", "unit": "oz", "type": "number", "calc": "IF({Is_Ore}=1,{Contained_oz},0)", "sql": "CASE WHEN material LIKE 'Ore%' THEN tonnes_t * COALESCE(grade_gpt, 0) / {TROY} ELSE 0 END", "desc": "Contained oz for ore rows only"}
      ]
    },
    {
      "table": "tblMiningDaily",
      "sheet": "Mining_Daily",
      "title": "Mining daily totals, one row per day",
      "owner": "Mining (production engineer)",
      "prefill": "dates",
      "fields": [
        {"name": "Date", "label": "Date", "type": "date", "key": true},
        {"name": "Drill_m", "label": "Drilled", "unit": "m", "type": "number", "min": 0, "max": 20000, "desc": "Production drill metres"},
        {"name": "GC_Drill_m", "label": "Grade control drilled", "unit": "m", "type": "number", "min": 0, "max": 5000, "desc": "Grade control RC drill metres"},
        {"name": "Blast_t", "label": "Blasted", "unit": "t", "type": "number", "min": 0, "max": 500000, "desc": "Tonnes blasted"},
        {"name": "Explosives_kg", "label": "Explosives", "unit": "kg", "type": "number", "min": 0, "max": 200000, "desc": "Explosives consumed"},
        {"name": "Diesel_L", "label": "Diesel", "unit": "L", "type": "number", "min": 0, "max": 200000, "desc": "Mining fleet diesel"},
        {"name": "Dewatering_m3", "label": "Dewatering", "unit": "m3", "type": "number", "min": 0, "max": 500000, "desc": "Pit water pumped"},
        {"name": "Comments", "label": "Comments", "type": "text"},
        {"name": "Powder_Factor_kgpt", "label": "Powder factor", "unit": "kg/t", "type": "number", "calc": "IF(AND({Blast_t}<>\"\",{Blast_t}>0),{Explosives_kg}/{Blast_t},\"\")", "sql": "CASE WHEN blast_t > 0 THEN COALESCE(explosives_kg, 0) / blast_t END", "desc": "Explosives kg per tonne blasted"}
      ]
    },
    {
      "table": "tblFleet",
      "sheet": "Fleet",
      "title": "Mining fleet time model, one row per equipment class per day",
      "owner": "Mining (dispatch / maintenance planner)",
      "rows": 3000,
      "unique": [["Date", "Equipment_Class"]],
      "fields": [
        {"name": "Date", "label": "Date", "type": "date", "required": true},
        {"name": "Equipment_Class", "label": "Equipment class", "type": "list", "list": "Equipment_Class", "required": true},
        {"name": "Units", "label": "Units in fleet", "type": "int", "min": 0, "max": 200, "required": true, "desc": "Number of machines of this class on site"},
        {"name": "Down_h", "label": "Down hours", "unit": "h", "type": "number", "min": 0, "max": 4800, "desc": "Sum of maintenance and breakdown hours across the class"},
        {"name": "Operating_h", "label": "Operating hours", "unit": "h", "type": "number", "min": 0, "max": 4800, "desc": "Sum of SMU engine hours across the class (one basis for every class)"},
        {"name": "Comments", "label": "Comments", "type": "text"},
        {"name": "Calendar_h", "label": "Calendar hours", "unit": "h", "type": "number", "calc": "IF({Units}=\"\",\"\",{Units}*{HOURS_PER_DAY})", "sql": "units * 24", "desc": "Units x 24"},
        {"name": "Available_h", "label": "Available hours", "unit": "h", "type": "number", "calc": "IF({Units}=\"\",\"\",{Calendar_h}-{Down_h})", "sql": "units * 24 - COALESCE(down_h, 0)", "desc": "Calendar minus down"},
        {"name": "Availability_pct", "label": "Availability", "unit": "%", "type": "pct", "calc": "IF(AND({Units}<>\"\",{Calendar_h}>0),{Available_h}/{Calendar_h},\"\")", "sql": "CASE WHEN units > 0 THEN (units * 24 - COALESCE(down_h, 0)) / (units * 24.0) END", "desc": "Available over calendar"},
        {"name": "Utilisation_pct", "label": "Utilisation", "unit": "%", "type": "pct", "calc": "IF(AND({Units}<>\"\",{Available_h}>0),{Operating_h}/{Available_h},\"\")", "sql": "CASE WHEN (units * 24 - COALESCE(down_h, 0)) > 0 THEN COALESCE(operating_h, 0) / (units * 24 - COALESCE(down_h, 0)) END", "desc": "Operating over available"}
      ]
    },
    {
      "table": "tblSafety",
      "sheet": "Safety",
      "title": "Safety, health, environment and community, one row per day",
      "owner": "HSE",
      "prefill": "dates",
      "fields": [
        {"name": "Date", "label": "Date", "type": "date", "key": true},
        {"name": "Hours_Employees", "label": "Hours worked, employees", "unit": "h", "type": "number", "min": 0, "max": 50000},
        {"name": "Hours_Contractors", "label": "Hours worked, contractors", "unit": "h", "type": "number", "min": 0, "max": 50000},
        {"name": "Fatality", "label": "Fatalities", "type": "int", "min": 0, "max": 50},
        {"name": "LTI", "label": "Lost time injuries", "type": "int", "min": 0, "max": 50},
        {"name": "RWI", "label": "Restricted work injuries", "type": "int", "min": 0, "max": 50},
        {"name": "MTI", "label": "Medical treatment injuries", "type": "int", "min": 0, "max": 50},
        {"name": "FAI", "label": "First aid injuries", "type": "int", "min": 0, "max": 50},
        {"name": "Near_Miss", "label": "Near misses", "type": "int", "min": 0, "max": 500},
        {"name": "Hazard_Reports", "label": "Hazard reports", "type": "int", "min": 0, "max": 5000},
        {"name": "HPI", "label": "High potential incidents", "type": "int", "min": 0, "max": 50},
        {"name": "Env_Incidents", "label": "Environmental incidents", "type": "int", "min": 0, "max": 50},
        {"name": "Community_Incidents", "label": "Community incidents or grievances", "type": "int", "min": 0, "max": 50},
        {"name": "Vehicle_Incidents", "label": "Vehicle incidents", "type": "int", "min": 0, "max": 50},
        {"name": "Toolbox_Talks", "label": "Toolbox talks", "type": "int", "min": 0, "max": 500},
        {"name": "Inspections", "label": "Inspections and audits", "type": "int", "min": 0, "max": 500},
        {"name": "Comments", "label": "Comments", "type": "text"},
        {"name": "Hours_Total", "label": "Hours worked, total", "unit": "h", "type": "number", "calc": "IF(AND({Hours_Employees}=\"\",{Hours_Contractors}=\"\"),\"\",N({Hours_Employees})+N({Hours_Contractors}))", "sql": "COALESCE(hours_employees, 0) + COALESCE(hours_contractors, 0)"},
        {"name": "Recordables", "label": "Recordable injuries", "type": "int", "calc": "IF(AND({Fatality}=\"\",{LTI}=\"\",{RWI}=\"\",{MTI}=\"\"),\"\",N({Fatality})+N({LTI})+N({RWI})+N({MTI}))", "sql": "COALESCE(fatality,0) + COALESCE(lti,0) + COALESCE(rwi,0) + COALESCE(mti,0)", "desc": "Fatalities + LTI + RWI + MTI"}
      ]
    },
    {
      "table": "tblSafetyHistory",
      "sheet": "Config",
      "title": "Monthly safety history for months without daily Safety rows (prior year, and this year before go-live), for rolling 12-month rates",
      "owner": "HSE",
      "prefill": "prior_months",
      "fields": [
        {"name": "Month", "label": "Month", "type": "date", "key": true, "desc": "First day of the month"},
        {"name": "Hours_Total", "label": "Hours worked", "unit": "h", "type": "number", "min": 0, "max": 2000000},
        {"name": "Recordables", "label": "Recordable injuries", "type": "int", "min": 0, "max": 100},
        {"name": "LTI", "label": "Lost time injuries", "type": "int", "min": 0, "max": 100},
        {"name": "Use", "label": "Used in rates", "type": "int", "calc": "IF({Month}=\"\",\"\",IF(COUNTIFS(tblSafety[Date],\">=\"&{Month},tblSafety[Date],\"<=\"&EOMONTH({Month},0),tblSafety[Hours_Total],\">0\")=0,1,0))", "desc": "1 when the Safety sheet has no hours for the month, so this row supplies it; 0 when daily rows exist"}
      ]
    },
    {
      "table": "tblGold",
      "sheet": "Gold",
      "title": "Gold shipments and sales",
      "owner": "Finance / gold room",
      "rows": 400,
      "fields": [
        {"name": "Date", "label": "Date", "type": "date", "required": true},
        {"name": "Type", "label": "Type", "type": "list", "list": "Gold_Movement_Type", "required": true, "desc": "Shipment (dore leaves site) or Sale (refinery outturn sold)"},
        {"name": "Dore_kg", "label": "Dore weight", "unit": "kg", "type": "number", "min": 0, "max": 5000},
        {"name": "Gold_oz", "label": "Gold", "unit": "oz", "type": "number", "min": 0, "max": 100000, "required": true},
        {"name": "Silver_oz", "label": "Silver", "unit": "oz", "type": "number", "min": 0, "max": 100000},
        {"name": "Reference", "label": "Reference", "type": "text", "desc": "Shipment or sale reference"},
        {"name": "Comments", "label": "Comments", "type": "text"}
      ]
    },
    {
      "table": "tblBudget",
      "sheet": "Budget",
      "title": "Monthly budget (or latest forecast), one row per month",
      "owner": "Technical services / finance",
      "prefill": "months",
      "fields": [
        {"name": "Month", "label": "Month", "type": "date", "key": true, "desc": "First day of the month"},
        {"name": "Ore_Mined_t", "label": "Ore mined", "unit": "t", "type": "number", "min": 0, "max": 5000000},
        {"name": "Ore_Grade_gpt", "label": "Ore grade mined", "unit": "g/t", "type": "number", "min": 0, "max": 50},
        {"name": "Waste_t", "label": "Waste mined", "unit": "t", "type": "number", "min": 0, "max": 20000000},
        {"name": "Milled_t", "label": "Milled", "unit": "t", "type": "number", "min": 0, "max": 5000000},
        {"name": "Head_Grade_gpt", "label": "Head grade", "unit": "g/t", "type": "number", "min": 0, "max": 50},
        {"name": "Recovery_pct", "label": "Recovery", "unit": "%", "type": "pct", "min": 0, "max": 1},
        {"name": "Gold_Poured_oz", "label": "Gold poured", "unit": "oz", "type": "number", "min": 0, "max": 500000},
        {"name": "Hours_Worked", "label": "Hours worked", "unit": "h", "type": "number", "min": 0, "max": 2000000},
        {"name": "Drill_m", "label": "Drilled", "unit": "m", "type": "number", "min": 0, "max": 500000},
        {"name": "Diesel_L", "label": "Mining diesel", "unit": "L", "type": "number", "min": 0, "max": 10000000},
        {"name": "Days", "label": "Days in month", "type": "int", "calc": "DAY(EOMONTH({Month},0))", "sql": "EXTRACT(DAY FROM (date_trunc('month', month::timestamp) + INTERVAL '1 month - 1 day'))::integer", "sql_postgres": "EXTRACT(DAY FROM (date_trunc('month', month::timestamp) + INTERVAL '1 month - 1 day'))::integer", "sql_sqlite": "CAST(strftime('%d', date(month, 'start of month', '+1 month', '-1 day')) AS INTEGER)"},
        {"name": "Ore_oz", "label": "Ore gold mined", "unit": "oz", "type": "number", "calc": "N({Ore_Mined_t})*N({Ore_Grade_gpt})/{TROY}", "sql": "COALESCE(ore_mined_t,0) * COALESCE(ore_grade_gpt,0) / {TROY}"},
        {"name": "TMM_t", "label": "Total material moved", "unit": "t", "type": "number", "calc": "N({Ore_Mined_t})+N({Waste_t})", "sql": "COALESCE(ore_mined_t,0) + COALESCE(waste_t,0)"},
        {"name": "Feed_oz", "label": "Contained gold in feed", "unit": "oz", "type": "number", "calc": "N({Milled_t})*N({Head_Grade_gpt})/{TROY}", "sql": "COALESCE(milled_t,0) * COALESCE(head_grade_gpt,0) / {TROY}"},
        {"name": "Recovered_oz", "label": "Gold recovered", "unit": "oz", "type": "number", "calc": "{Feed_oz}*N({Recovery_pct})", "sql": "COALESCE(milled_t,0) * COALESCE(head_grade_gpt,0) / {TROY} * COALESCE(recovery_pct,0)"}
      ]
    },
    {
      "table": "tblCommentary",
      "sheet": "Commentary",
      "title": "Daily report narrative, one row per comment",
      "owner": "All departments; compiled by the production engineer",
      "rows": 3000,
      "fields": [
        {"name": "Date", "label": "Date", "type": "date", "required": true},
        {"name": "Area", "label": "Area", "type": "list", "list": "Area", "required": true},
        {"name": "Comment", "label": "Comment", "type": "text", "required": true, "desc": "One or two sentences; appears on the daily report"},
        {"name": "Author", "label": "Author", "type": "text"},
        {"name": "Seq_Day", "label": "Sequence in day", "type": "int", "calc": "IF({Date}=\"\",\"\",COUNTIFS(INDEX(tblCommentary[Date],1):{Date},{Date}))", "desc": "Running number of comments for the same day (report helper)"},
        {"name": "Key_Day", "label": "Day key", "type": "text", "calc": "IF({Date}=\"\",\"\",{Date}&\"|\"&{Seq_Day})", "desc": "Report lookup key"}
      ]
    }
  ]
}
'@
Write-TextFile (Join-Path $Root "01_Tools\mps\schema\mps_schema.json") $content_mps_schema_mps_schema_json $false
$content_mps_README_md = @'
# mps

Generator, validator and SQL builder for the Mako Production System (MPS) workbook. Everything about the data model lives in `schema/mps_schema.json`; the scripts render it. See `docs/MPS_Design.md` for the design and `docs/data_dictionary.md` for the field-level dictionary. On the site PC the same files sit in `C:\MakoPS\01_Tools\mps` and are driven by the PowerShell scripts described in `tools/README.md`.

## Requirements

- Python 3.11 or later with `openpyxl` 3.1 (`pip install openpyxl`).
- For the validator only: LibreOffice 24 Calc (headless recalculation) and `oletools` (`pip install oletools`) for the VBA check in `tools/xl_inspect.py`.

Excel 2016 or later opens the workbook. Functions that Excel introduced after 2010 are written with the `_xlfn.` prefix (MAXIFS, ISOWEEKNUM) so both Excel and LibreOffice evaluate them; XLOOKUP is not used.

## Build the workbook

From the repository root:

```bash
# blank workbook for the site, dates prefilled for the year, no data
python3 mps/build_workbook.py --out dist/MPS_2026.xlsx

# DEMO workbook: 74 days of deterministic synthetic data, title marked [DEMO DATA]
python3 mps/build_workbook.py --out dist/MPS_2026_DEMO.xlsx --sample-days 74

# another year
python3 mps/build_workbook.py --out dist/MPS_2027.xlsx --year 2027

# fix the report date instead of the default (yesterday or last data date); used by tests
python3 mps/build_workbook.py --out dist/demo.xlsx --sample-days 74 --report-date 2026-02-10
```

The last line of the output states the sheet, table, name and formula counts (17 sheets, 20 tables, 282 names, about 146,000 formulas for 2026). The file is saved with full recalculation on load, so Excel computes everything the first time it is opened; save it once from Excel before distributing so cached values exist for readers that do not recalculate.

On site, `C:\MakoPS\01_Tools\Build-MPS.ps1` runs these two builds (blank and DEMO) into `C:\MakoPS\04_MPS`, finding or installing Python first. `dist/` holds the latest built pair (blank and DEMO), committed so they can be downloaded from the branch.

## Run the validator

```bash
python3 mps/validate_workbook.py            # fresh temp folder, deleted afterwards
python3 mps/validate_workbook.py --keep --workdir /tmp/mps_check
python3 mps/validate_workbook.py --soffice "C:\Program Files\LibreOffice\program\soffice.exe"
```

It builds three workbooks (DEMO with the report date fixed to 10 February, DEMO with the default report date, and blank), recalculates them with LibreOffice, recomputes the key figures in plain Python from the synthetic rows without using any workbook formula, and compares them with Calc_Daily on 1 January, 10 February and the last day with data. It then checks Daily_Report, Monthly_Summary and Chart_Data against Calc_Daily, scans every sheet for error values, runs `tools/xl_inspect.py` on the built file (no external links, no VBA, no dangling references) and confirms the blank workbook recalculates without a single error. One line per check, about 270 checks; exit code 1 on any failure. Allow about a minute.

Run it after every change to the schema or the generator.

## Change the schema and regenerate

Edit `schema/mps_schema.json`, then rebuild, validate, regenerate the DDL and repack the bootstrap that carries the schema and generator to site:

```bash
python3 mps/build_workbook.py --out dist/MPS_2026.xlsx
python3 mps/validate_workbook.py
python3 mps/build_ddl.py
python3 tools/build_setup.py
```

What the schema holds:

| Key | Meaning |
|---|---|
| `lists` | Drop-down values. Each becomes a table on the Lists sheet, a dynamic named range `lst_<name>` and a `ref_<name>` table in SQL. |
| `reference_tables` | Sources, Destinations (sheet Lists) and Stockpiles (sheet Config) with seed rows. |
| `derived_lists` | Values used by reference table columns (Source_Type, Destination_Type). |
| `tables` | The nine input tables. `table` is the Excel table name, `sheet` the tab, `owner` the department, `prefill` one of `dates`, `months`, `prior_months`; otherwise `rows` gives the capacity. |
| `tables[].fields` | `name`, `label`, `type` (`date`, `text`, `list`, `number`, `int`, `pct`), `unit`, `min`, `max`, `required`, `key`, `list`, `desc`, `calc`, `sql`. |

A field with `calc` becomes a grey calculated column. Inside `calc`, `{Col}` is a this-row reference to another column of the same table and `{TROY}` is the troy ounce constant. Give the same field an `sql` expression (column names in lower case) so the DDL can generate it; a `calc` without `sql` stays Excel-only and is listed as such in the SQL header.

Rules of thumb:

- Adding a drop-down value does not need a rebuild: type it at the bottom of the list on the Lists sheet. Mirror it in the schema so the next build carries it.
- Adding or renaming a column, a table or a validation range needs a rebuild. The generator writes an empty workbook; move existing data across by pasting values into the matching columns of the new file, column by column if the order changed.
- Bump `version` in the schema for any change that alters a definition. It is written to the README sheet, the SQL headers and the dictionary.
- Never edit the built workbook's structure by hand; the next build would silently drop the change.
- `Setup-MakoPS.ps1` embeds the schema, the three scripts and this README, so commit the regenerated bootstrap together with the change (`python3 tools/build_setup.py --check` verifies it).

KPIs live in the `KPIS` list at the top of `build_workbook.py`, one `K(...)` entry each: `sum` (a SUMIFS over one table, with MTD and YTD), `ratio` (numerator and denominator built from other KPIs, recomputed per period), `custom` (an explicit Day formula, used for rolling rates and days since LTI) and `point` (a value at the date). `budget` names a tblBudget column; ratio budgets are derived when both sides are budgeted. `REPORT_LAYOUT` and `MONTHLY_KPIS` decide what appears on Daily_Report and Monthly_Summary. `build_ddl.py` translates plain SUMIFS and arithmetic automatically; anything else needs an entry in its `KPI_SQL` map.

## Build the DDL

```bash
python3 mps/build_ddl.py
```

No arguments; output is byte-identical on every run. It writes:

| File | Content |
|---|---|
| `mps/sql/schema_postgres.sql` | PostgreSQL 13 or later. Tables, CHECK constraints from min and max, foreign keys to the list tables, generated columns for calculated fields, `<table>_calc` views for lookups, `kpi_daily` and `kpi_monthly` views, seeds, `updated_at` triggers. |
| `mps/sql/schema_sqlite.sql` | The same for SQLite 3.31 or later, for local testing. |
| `docs/data_dictionary.md` | Field-level dictionary and the Excel to SQL mapping. |

Apply with `psql -v ON_ERROR_STOP=1 -1 -f mps/sql/schema_postgres.sql` (single transaction) or `sqlite3 mps.db < mps/sql/schema_sqlite.sql`. Both scripts are safe to re-run on an existing database; structural changes to populated tables still need a migration.

## Folder layout

| Path | Purpose |
|---|---|
| `mps/schema/mps_schema.json` | The data model. Single source for the workbook, the SQL and the dictionary. |
| `mps/build_workbook.py` | Workbook generator and KPI catalogue. |
| `mps/validate_workbook.py` | Build, recalculate, verify. |
| `mps/build_ddl.py` | SQL and dictionary generator. |
| `mps/sql/` | Generated DDL, committed so reviewers can read it without running anything. |
| `docs/MPS_Design.md` | Design note. |
| `docs/data_dictionary.md` | Generated dictionary. |
| `tools/` | Workbook inspector, the site PowerShell scripts and the bootstrap builder. See `tools/README.md`. |
| `Setup-MakoPS.ps1` | Generated bootstrap; carries the `mps` files to the site PC. |
| `dist/` | Latest built blank and DEMO workbooks. |

## C:\MakoPS layout on site

`Setup-MakoPS.ps1` creates and fills the layout, then runs the inspection and builds the workbooks.

| Folder | Content |
|---|---|
| `01_Tools` | `PythonEnv.ps1`, `Invoke-XlInspect.ps1`, `Build-MPS.ps1`, `Export-DailyReport.ps1`, `xl_inspect.py`, `README.md`. |
| `01_Tools\mps` | `build_workbook.py`, `build_ddl.py`, `validate_workbook.py`, `schema\mps_schema.json` and this README. Rebuilt from the repository by the bootstrap; do not edit here. |
| `02_Source\2026` | Working copies of the legacy PMC workbooks. Originals on X: are not touched. |
| `03_Inspection` | `inspection_report.md` and one JSON per legacy workbook. |
| `04_MPS` | `MPS_2026.xlsx`, the live workbook and the only file anyone edits, and `MPS_2026_DEMO.xlsx` for training. `Export-DailyReport.ps1` writes the PDFs to `04_MPS\Reports\Mako_Daily_Report_yyyy-MM-dd.pdf`. Keep dated backup copies in `04_MPS\Backup`. |

Re-running `Build-MPS.ps1` overwrites both workbooks, so copy the live file to `Backup` first and paste its table bodies into the new build. Re-running the inspector on `04_MPS` (`python xl_inspect.py C:\MakoPS\04_MPS -o C:\MakoPS\03_Inspection\mps`) is the quickest proof that the live workbook still has no external links, no VBA and no error values.
'@
Write-TextFile (Join-Path $Root "01_Tools\mps\README.md") $content_mps_README_md $false

# ---------------------------------------------------------------- copy workbooks
$sourceCopy = Join-Path $Root "02_Source\2026"
if (-not $SkipCopy) {
    if (Test-Path $Source) {
        $files = Get-ChildItem -Path $Source -File | Where-Object {
            $_.Extension -match '^\.(xlsx|xlsm|xlsb|xls|xltm|xltx)$' -and $_.Name -notlike '~$*'
        }
        if ($files.Count -eq 0) {
            Write-Warning "No Excel files found in $Source"
        } else {
            Write-Host "Copying $($files.Count) workbook(s) from $Source to $sourceCopy"
            foreach ($f in $files) {
                Copy-Item -LiteralPath $f.FullName -Destination (Join-Path $sourceCopy $f.Name) -Force
                Write-Host ("  {0}  ({1:N1} MB)" -f $f.Name, ($f.Length / 1MB))
            }
        }
    } else {
        Write-Warning "Source folder not reachable: $Source. Copy the workbooks into $sourceCopy manually, then run 01_Tools\Invoke-XlInspect.ps1."
        if ($elevated) { Write-Warning "The session is elevated, which hides mapped drives such as X:. Re-run from a normal PowerShell window." }
    }
}

# ---------------------------------------------------------------- python (resolved once for both scripts below)
$pyOk = $false
if (-not ($SkipRun -and $SkipBuild)) {
    Write-Host ""
    . (Join-Path $Root "01_Tools\PythonEnv.ps1")
    try {
        $py = Resolve-Python
        Write-Host "Using Python: $py"
        $pyOk = $true
    } catch {
        Write-Warning $_.Exception.Message
        Write-Warning "Inspection and workbook build skipped. Install Python 3, then run 01_Tools\Invoke-XlInspect.ps1 and 01_Tools\Build-MPS.ps1 from $Root."
    }
}

# ---------------------------------------------------------------- run inspection
if (-not $SkipRun -and $pyOk) {
    $runner = Join-Path $Root "01_Tools\Invoke-XlInspect.ps1"
    $books = @(Get-ChildItem -Path $sourceCopy -File -ErrorAction SilentlyContinue | Where-Object {
        $_.Extension -match '^\.(xlsx|xlsm|xlsb|xls|xltm|xltx)$' -and $_.Name -notlike '~$*'
    })
    if ($books.Count -eq 0) {
        Write-Warning "No workbooks in $sourceCopy, skipping the inspection. Copy them there and run $runner later."
    } else {
        Write-Host ""
        Write-Host "Running inspection"
        try {
            & $runner -Source $sourceCopy -Out (Join-Path $Root "03_Inspection") -NoInstall
        } catch {
            Write-Warning "Inspection failed: $($_.Exception.Message). Fix the cause and run $runner again."
        }
        $report = Join-Path $Root "03_Inspection\inspection_report.md"
        if (Test-Path $report) {
            Write-Host ""
            Write-Host "Inspection report: $report"
        }
    }
}

# ---------------------------------------------------------------- build MPS workbooks
if (-not $SkipBuild -and $pyOk) {
    $builder = Join-Path $Root "01_Tools\Build-MPS.ps1"
    Write-Host ""
    Write-Host "Building the MPS workbooks for $Year"
    try {
        & $builder -Root $Root -Year $Year -NoInstall -NoOpen
    } catch {
        Write-Warning "Build failed: $($_.Exception.Message). Fix the cause and run $builder again."
    }
}

Write-Host ""
Write-Host "Setup finished. Tools are in $(Join-Path $Root '01_Tools'); see README.md there."
# one Explorer window on the root, after everything has been written
try { Start-Process explorer.exe $Root } catch { }
