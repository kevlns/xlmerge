"""Safe, Excel-backed formula materialisation for the conflict resolver.

This module deliberately has a small trust boundary.  Excel is only ever given a
temporary copy, with links and macros disabled.  The result produced by Excel is
read back as cached values and written into a copy of the original package with
openpyxl.  We do not use ``data_only`` as a source workbook: a stale cache is not
an acceptable answer for a conflict resolver.
"""

from __future__ import annotations

import json
import math
import os
import re
import shutil
import subprocess
import tempfile
import time
import zipfile
from datetime import date, datetime, time as datetime_time
from pathlib import Path
from typing import Any, Iterable
from xml.etree import ElementTree

from openpyxl import load_workbook
from openpyxl.cell.rich_text import CellRichText, InlineFont, TextBlock
from openpyxl.formula import Tokenizer
from openpyxl.utils.cell import range_boundaries


class FormulaValueError(RuntimeError):
    """Raised when a workbook cannot be evaluated and verified safely."""


DEFAULT_TIMEOUT_SECONDS = 120
_ERROR_VALUES = {"#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A", "#GETTING_DATA"}
_VOLATILE_FUNCTIONS = {
    "NOW", "TODAY", "RAND", "RANDBETWEEN", "RANDARRAY", "RTD", "OFFSET", "INDIRECT",
    "CELL", "INFO",
}
_DYNAMIC_FUNCTIONS = {
    "FILTER", "SORT", "SORTBY", "UNIQUE", "SEQUENCE", "RANDARRAY", "TRANSPOSE",
    "TOCOL", "TOROW", "TAKE", "DROP", "CHOOSECOLS", "CHOOSEROWS", "EXPAND",
    "VSTACK", "HSTACK", "WRAPROWS", "WRAPCOLS", "XLOOKUP", "XMATCH", "LAMBDA",
}

# Function names are checked before opening Excel.  This is intentionally a
# broad list of ordinary worksheet functions, but unknown calls are rejected so
# a disabled VBA UDF can never be silently accepted merely because it returned a
# plausible value in a cached workbook.
_BUILTIN_FUNCTIONS = {
    name
    for name in """
    ABS ACOS ACOSH ACOT ACOTH AGGREGATE ARABIC ASIN ASINH ATAN ATAN2 ATANH BASE
    CEILING CEILING.MATH CEILING.PRECISE COMBIN COMBINA COS COSH COT COTH CSC CSCH
    DECIMAL DEGREES EVEN EXP FACT FACTDOUBLE FLOOR FLOOR.MATH FLOOR.PRECISE GCD
    INT ISO.CEILING LCM LN LOG LOG10 MDETERM MINVERSE MMULT MOD MROUND MULTINOMIAL
    MUNIT ODD PI POWER PRODUCT QUOTIENT RADIANS ROMAN ROUND ROUNDDOWN ROUNDUP SEC
    SECH SIGN SIN SINH SQRT SUM SUMIF SUMIFS SUMPRODUCT SUMSQ TAN TANH TRUNC
    ACCRINT ACCRINTM AMORDEGRC AMORLINC COUPDAYBS COUPDAYS COUPDAYSNC COUPNCD
    COUPNUM COUPPCD CUMIPMT CUMPRINC DB DDB DISC DOLLARDE DOLLARFR DURATION EFFECT
    FV FVSCHEDULE INTRATE IPMT IRR ISPMT MDURATION MIRR NOMINAL NPER NPV PDURATION
    PMT PPMT PRICE PRICEDISC PRICEMAT PV RATE RECEIVED RRI SLN SYD TBILLEQ TBILLPRICE
    TBILLYIELD VDB XIRR XNPV YIELD YIELDDISC YIELDMAT
    AND FALSE IF IFERROR IFNA NOT OR TRUE XOR SWITCH IFS CHOOSE
    ADDRESS AREAS COLUMN COLUMNS FORMULATEXT HLOOKUP HYPERLINK INDEX INDIRECT
    LOOKUP MATCH ROW ROWS TRANSPOSE VLOOKUP OFFSET
    CELL ERROR.TYPE INFO ISBLANK ISERR ISERROR ISEVEN ISFORMULA ISLOGICAL ISNA
    ISNONTEXT ISNUMBER ISODD ISREF ISTEXT N TYPE NA
    CONCAT CONCATENATE EXACT FIND FINDB LEFT LEFTB LEN LENB LOWER MID MIDB NUMBERVALUE
    PHONETIC PROPER REPLACE REPLACEB REPT RIGHT RIGHTB SEARCH SEARCHB SUBSTITUTE T
    TEXT TEXTAFTER TEXTBEFORE TEXTJOIN TEXTSPLIT TRIM UNICHAR UNICODE UPPER VALUE
    ASC CHAR CLEAN CODE DBCS FIXED JIS LOWER UPPER
    DATE DATEVALUE DAY DAYS DAYS360 EDATE EOMONTH HOUR ISOWEEKNUM MINUTE MONTH
    NETWORKDAYS NETWORKDAYS.INTL SECOND TIME TIMEVALUE WEEKDAY WEEKNUM WORKDAY
    WORKDAY.INTL YEAR YEARFRAC
    COUNT COUNTA COUNTBLANK COUNTIF COUNTIFS MAX MAXIFS MIN MINIFS AVERAGE
    AVERAGEIF AVERAGEIFS MEDIAN MODE.MULT MODE.SNGL LARGE SMALL RANK.EQ RANK.AVG
    PERCENTILE.EXC PERCENTILE.INC PERCENTRANK.EXC PERCENTRANK.INC QUARTILE.EXC
    QUARTILE.INC STANDARDIZE STDEV.P STDEV.S STDEVA STDEVPA VAR.P VAR.S VARA VARPA
    CORREL COVARIANCE.P COVARIANCE.S GEOMEAN HARMEAN KURT SKEW SKEW.P SLOPE
    INTERCEPT FORECAST FORECAST.LINEAR FREQUENCY GROWTH LINEST LOGEST PEARSON RSQ
    TREND
    DEC2BIN DEC2HEX DEC2OCT HEX2BIN HEX2DEC HEX2OCT OCT2BIN OCT2DEC OCT2HEX
    BITAND BITLSHIFT BITOR BITRSHIFT BITXOR CONVERT DELTA ERF ERF.PRECISE ERFC
    ERFC.PRECISE GESTEP IMSUM IMABS IMAGINARY IMARGUMENT IMCONJUGATE IMCOS IMCOSH
    IMCOT IMCSC IMCSCH IMDIV IMEXP IMLN IMLOG10 IMLOG2 IMPOWER IMPRODUCT IMREAL
    IMSEC IMSECH IMSIN IMSINH IMSQRT IMSUB IMTAN IMTANH COMPLEX
    CUBEMEMBER CUBEVALUE GETPIVOTDATA
    LET LAMBDA MAP REDUCE SCAN MAKEARRAY BYROW BYCOL
    """.split()
}

_FUNCTION_CALL_RE = re.compile(r"(?<![A-Za-z0-9_.])([A-Za-z_][A-Za-z0-9_.]*)\s*\(")
_EXTERNAL_REF_RE = re.compile(r"\[[^\]]+\][^!]*!")
_DYNAMIC_SPILL_RE = re.compile(r"(?:[A-Z]{1,3}\$?\d+|\))#(?=$|[\s,;)])", re.IGNORECASE)


def _powershell_executable() -> str:
    # Windows PowerShell is part of the supported host.  Do not fall back to a
    # different Python/COM package; pywin32 is intentionally not required.
    return shutil.which("powershell.exe") or shutil.which("powershell") or "powershell.exe"


def _formula_text(value: Any) -> str | None:
    if isinstance(value, str) and value.startswith("="):
        return value
    return None


def _without_string_literals(formula: str) -> str:
    """Blank double-quoted Excel string literals before syntax/security scans."""
    out: list[str] = []
    quoted = False
    index = 0
    while index < len(formula):
        char = formula[index]
        if char == '"':
            if quoted and index + 1 < len(formula) and formula[index + 1] == '"':
                out.extend("  ")
                index += 2
                continue
            quoted = not quoted
            out.append(" ")
        else:
            out.append(" " if quoted else char)
        index += 1
    return "".join(out)


def _iter_formula_cells(workbook: Any) -> Iterable[tuple[Any, Any]]:
    for worksheet in workbook.worksheets:
        for row in worksheet.iter_rows():
            for cell in row:
                # A literal string beginning with '=' is data_type='s' and is
                # deliberately not a formula.  Only Excel's formula cell type
                # (including openpyxl ArrayFormula objects) enters this path.
                if cell.data_type == "f":
                    yield worksheet, cell


def _scan_package(path: Path) -> None:
    """Reject package features that can cause I/O or code execution in Excel."""
    try:
        with zipfile.ZipFile(path) as archive:
            archive_names = {name.replace("\\", "/").lower(): name for name in archive.namelist()}
            names = set(archive_names)
            dangerous_parts = (
                "xl/externallinks/", "xl/connections.xml", "xl/querytables/",
                "xl/ddelinks/", "xl/macrosheets/", "xl/dialogsheets/",
            )
            for part in dangerous_parts:
                if any(name.startswith(part) or name == part.rstrip("/") for name in names):
                    raise FormulaValueError(f"Workbook contains unsupported external or executable part: {part}")

            for name in names:
                if not name.endswith(".xml") and not name.endswith(".rels"):
                    continue
                raw = archive.read(archive_names[name])
                lowered = raw.lower()
                if b"<ddelink" in lowered or b"<olelink" in lowered:
                    raise FormulaValueError(f"Workbook contains DDE/OLE link data: {name}")
                if name.endswith(".rels") and b'targetmode="external"' in lowered:
                    # External hyperlinks are harmless, but an external target
                    # in workbook/externalLink relationships is not.  Reject
                    # only relationship files whose type names indicate a link
                    # or connection, avoiding false positives for user URLs.
                    if b"externallink" in lowered or b"externalconnection" in lowered:
                        raise FormulaValueError(f"Workbook contains an external relationship: {name}")
            # Shared and array formulas are represented by attributes that are
            # not faithfully round-trippable through openpyxl.  Refuse the
            # whole workbook rather than replacing only the anchor and leaving
            # a silently wrong spill/shared range.
            for name in names:
                if not name.startswith("xl/worksheets/") or not name.endswith(".xml"):
                    continue
                root = ElementTree.fromstring(archive.read(archive_names[name]))
                for formula_node in root.iter():
                    if formula_node.tag.rsplit("}", 1)[-1] != "f":
                        continue
                    attrs = {key.rsplit("}", 1)[-1]: value for key, value in formula_node.attrib.items()}
                    kind = attrs.get("t", "").lower()
                    if kind in {"shared", "array", "datatable"} or "ref" in attrs:
                        raise FormulaValueError(f"Unsupported shared/array formula in {name}")
                    if formula_node.text is None and kind:
                        raise FormulaValueError(f"Unsupported formula without body in {name}")
    except ElementTree.ParseError as exc:
        raise FormulaValueError(f"Cannot inspect workbook XML safely: {path}") from exc
    except (OSError, zipfile.BadZipFile) as exc:
        raise FormulaValueError(f"Cannot inspect workbook package: {path}") from exc


def _scan_formulas(workbook: Any) -> list[tuple[str, str, str]]:
    formulas: list[tuple[str, str, str]] = []
    for worksheet, cell in _iter_formula_cells(workbook):
        formula = cell.value.text if hasattr(cell.value, "text") else cell.value
        if not isinstance(formula, str):
            raise FormulaValueError(f"Unsupported formula object at {worksheet.title}!{cell.coordinate}")
        if not formula.startswith("="):
            formula = "=" + formula
        scan_formula = _without_string_literals(formula)
        upper = scan_formula.upper()
        if _EXTERNAL_REF_RE.search(scan_formula) or re.search(r"\bDDE\s*\(", scan_formula, re.IGNORECASE):
            raise FormulaValueError(f"External workbook/DDE reference at {worksheet.title}!{cell.coordinate}")
        if any(token in upper for token in ("_XLFN.", "_XLWS.")):
            raise FormulaValueError(f"Unsupported modern/dynamic formula at {worksheet.title}!{cell.coordinate}")
        if _DYNAMIC_SPILL_RE.search(scan_formula) or "@" in scan_formula:
            raise FormulaValueError(f"Dynamic/implicit-array formula at {worksheet.title}!{cell.coordinate}")
        if any(re.search(rf"(?<![A-Z0-9_.]){re.escape(name)}\s*\(", upper) for name in _VOLATILE_FUNCTIONS):
            raise FormulaValueError(f"Volatile formula is not reproducible at {worksheet.title}!{cell.coordinate}")
        if "#REF!" in upper:
            raise FormulaValueError(f"Broken formula reference at {worksheet.title}!{cell.coordinate}")
        for match in _FUNCTION_CALL_RE.finditer(scan_formula):
            name = match.group(1).upper()
            if name in _DYNAMIC_FUNCTIONS:
                raise FormulaValueError(f"Dynamic-array formula at {worksheet.title}!{cell.coordinate}")
            if name.startswith("_XLFN.") or name.startswith("_XLWS.") or name not in _BUILTIN_FUNCTIONS:
                raise FormulaValueError(f"Unknown/UDF formula function {name} at {worksheet.title}!{cell.coordinate}")
        formulas.append((worksheet.title, cell.coordinate, formula))
    _detect_formula_cycles(workbook)
    return formulas


def _detect_formula_cycles(workbook: Any) -> None:
    """Fail closed on cycles or references we cannot statically account for."""
    nodes = {(ws.title, cell.coordinate): cell for ws, cell in _iter_formula_cells(workbook)}
    graph = {}
    sheet_names = {ws.title.casefold(): ws.title for ws in workbook.worksheets}
    global_names = {name.casefold(): value for name, value in workbook.defined_names.items()}
    local_names = {ws.title: {name.casefold(): value for name, value in ws.defined_names.items()}
                   for ws in workbook.worksheets}
    by_sheet = {}
    for key, cell in nodes.items():
        by_sheet.setdefault(key[0], []).append((key, cell.row, cell.column))
    ref_re = re.compile(r"^(?:\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?|\$?[A-Z]{1,3}:\$?[A-Z]{1,3}|\$?\d+:\$?\d+)$", re.I)

    def dependencies(formula, context_sheet, name_stack=()):
        try:
            tokens = Tokenizer(formula if formula.startswith('=') else '=' + formula).items
        except Exception as exc:
            raise FormulaValueError(f"Unable to parse formula dependencies in {context_sheet}") from exc
        result = set()
        for token in tokens:
            if token.type != "OPERAND" or token.subtype != "RANGE":
                continue
            raw = token.value.strip()
            if "!" in raw:
                sheet_text, raw_ref = raw.rsplit("!", 1)
                if ":" in sheet_text:
                    raise FormulaValueError(f"Unsupported 3-D formula reference: {raw}")
                if sheet_text.startswith("'") and sheet_text.endswith("'"):
                    sheet_text = sheet_text[1:-1].replace("''", "'")
                target_sheet = sheet_names.get(sheet_text.casefold())
                if target_sheet is None:
                    raise FormulaValueError(f"Unknown formula reference sheet: {raw}")
            else:
                raw_ref, target_sheet = raw, context_sheet
            if ref_re.fullmatch(raw_ref):
                # Relative references in defined names are anchored to Excel's
                # active-cell context, not necessarily the formula cell. Do
                # not guess that context when checking dependencies.
                if name_stack and not re.fullmatch(
                    r'(?:\$[A-Z]{1,3}\$\d+(?::\$[A-Z]{1,3}\$\d+)?|\$[A-Z]{1,3}:\$[A-Z]{1,3}|\$\d+:\$\d+)',
                    raw_ref, re.I,
                ):
                    raise FormulaValueError(f"Relative defined-name reference is unsupported: {raw}")
                min_col, min_row, max_col, max_row = range_boundaries(raw_ref.upper())
                min_col, min_row = min_col or 1, min_row or 1
                max_col, max_row = max_col or 16384, max_row or 1048576
                if not (1 <= min_col <= max_col <= 16384 and 1 <= min_row <= max_row <= 1048576):
                    raise FormulaValueError(f"Invalid formula range: {raw}")
                result.update(key for key, row, col in by_sheet.get(target_sheet, [])
                              if min_col <= col <= max_col and min_row <= row <= max_row)
            else:
                name_key = raw_ref.casefold()
                defined = local_names.get(target_sheet, {}).get(name_key) or global_names.get(name_key)
                if defined is None:
                    raise FormulaValueError(f"Unsupported or unknown formula reference: {raw}")
                identity = (target_sheet, name_key)
                if identity in name_stack or len(name_stack) >= 100:
                    raise FormulaValueError(f"Circular or excessively nested defined name: {raw}")
                result.update(dependencies(str(defined.attr_text or ''), target_sheet, (*name_stack, identity)))
        return result

    for node, cell in nodes.items():
        formula = cell.value.text if hasattr(cell.value, "text") else str(cell.value)
        graph[node] = dependencies(formula, node[0])
    # Iterative topological elimination supports long formula chains.
    reverse = {node: set() for node in graph}
    counts = {node: len(edges) for node, edges in graph.items()}
    for node, edges in graph.items():
        for edge in edges:
            reverse[edge].add(node)
    ready = [node for node, count in counts.items() if count == 0]
    visited = 0
    while ready:
        node = ready.pop()
        visited += 1
        for dependent in reverse[node]:
            counts[dependent] -= 1
            if counts[dependent] == 0:
                ready.append(dependent)
    if visited != len(graph):
        node = next(node for node, count in counts.items() if count)
        raise FormulaValueError(f"Circular formula dependency detected at {node[0]}!{node[1]}")


def _scan_defined_names(workbook: Any) -> None:
    """Reject names which smuggle an external link or executable reference."""
    all_names = [*workbook.defined_names.values(),
                 *(name for ws in workbook.worksheets for name in ws.defined_names.values())]
    for defined_name in all_names:
        text = str(getattr(defined_name, "attr_text", "") or "")
        scan_text = _without_string_literals(text)
        scan_upper = scan_text.upper()
        if _EXTERNAL_REF_RE.search(scan_text) or re.search(r"\bDDE\s*\(", scan_text, re.IGNORECASE) or "#REF!" in scan_upper:
            name = getattr(defined_name, "name", "<unnamed>")
            raise FormulaValueError(f"Defined name contains an unsafe reference: {name}")
        name = getattr(defined_name, "name", "<unnamed>")
        if any(re.search(rf"(?<![A-Z0-9_.]){re.escape(function)}\s*\(", scan_upper)
               for function in _VOLATILE_FUNCTIONS):
            raise FormulaValueError(f"Defined name contains a volatile formula: {name}")
        if _DYNAMIC_SPILL_RE.search(scan_text) or "@" in scan_text:
            raise FormulaValueError(f"Defined name contains a dynamic-array formula: {name}")
        for match in _FUNCTION_CALL_RE.finditer(scan_text):
            function = match.group(1).upper()
            if function in _DYNAMIC_FUNCTIONS or function.startswith(("_XLFN.", "_XLWS.")) \
                    or function not in _BUILTIN_FUNCTIONS:
                raise FormulaValueError(f"Defined name contains unknown/UDF function {function}: {name}")


def _read_state(path: Path) -> dict[str, Any]:
    try:
        text = path.read_text(encoding="utf-8-sig")
        value = json.loads(text)
        return value if isinstance(value, dict) else {}
    except (OSError, ValueError, json.JSONDecodeError):
        return {}


def _cleanup_excel_process(state_path: Path) -> None:
    """Stop only the Excel PID and start time recorded by our PS harness."""
    state = _read_state(state_path)
    pid = state.get("pid")
    start_filetime = state.get("startFileTimeUtc")
    if not isinstance(pid, int) or not isinstance(start_filetime, int):
        return
    helper = Path(__file__).with_name("recalculate_workbook.ps1")
    try:
        # Pass ownership data as JSON, never as PowerShell source/arguments.
        # The helper validates both PID and process creation time before stop.
        subprocess.run(
            [_powershell_executable(), "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass",
             "-File", str(helper), "-CleanupParametersPath", str(state_path)],
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=10, check=False,
        )
    except (OSError, subprocess.TimeoutExpired):
        pass


def _run_excel_recalculate(source: Path, result: Path, timeout_seconds: int) -> None:
    script = Path(__file__).with_name("recalculate_workbook.ps1")
    if not script.is_file():
        raise FormulaValueError(f"Excel recalculation helper is missing: {script}")
    with tempfile.TemporaryDirectory(prefix="xlmerge-excel-state-") as state_dir:
        state_path = Path(state_dir) / "state.json"
        params_path = Path(state_dir) / "params.json"
        params_path.write_text(json.dumps({
            "source": str(source), "result": str(result), "state": str(state_path),
            "timeoutSeconds": int(timeout_seconds),
        }), encoding="utf-8")
        command = [_powershell_executable(), "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass",
                   "-File", str(script), "-ParametersPath", str(params_path)]
        try:
            completed = subprocess.run(command, capture_output=True, text=True,
                                       encoding="utf-8", errors="replace",
                                       timeout=max(1, int(timeout_seconds) + 20), check=False)
        except subprocess.TimeoutExpired as exc:
            _cleanup_excel_process(state_path)
            raise FormulaValueError("Excel recalculation exceeded the hard timeout") from exc
        finally:
            # The helper normally closes Excel itself.  This also handles a
            # PowerShell error path without ever scanning/killing user Excel.
            if state_path.exists() and _read_state(state_path).get("status") not in {"completed", "failed"}:
                _cleanup_excel_process(state_path)
        if completed.returncode != 0:
            state = _read_state(state_path)
            detail = state.get("error") or completed.stderr.strip() or completed.stdout.strip()
            raise FormulaValueError(f"Excel could not safely calculate workbook: {detail or 'unknown error'}")
        if not result.is_file() or result.stat().st_size == 0:
            raise FormulaValueError("Excel reported success but did not export a result workbook")


def _set_literal(cell: Any, value: Any) -> None:
    """Assign a value without allowing a string beginning '=' to become a formula."""
    if isinstance(value, str):
        # openpyxl classifies strings such as '#N/A' as error cells and strings
        # beginning '=' as formulas.  Cached formula text is data, regardless
        # of its spelling, so force the serialized string type for all text.
        cell.data_type = "s"
        if value == "":
            # openpyxl normally omits an empty string cell entirely on write,
            # which reloads as None.  A zero-length rich-text run is a valid
            # inline string and round-trips as a typed empty string.
            cell._value = CellRichText(TextBlock(InlineFont(), ""))
        else:
            cell._value = value
    else:
        cell.value = value


def _validate_result(value: Any, worksheet: str, coordinate: str, data_type: str | None = None) -> None:
    # Use the serialized cell type, not the displayed text: a formula may
    # intentionally return the literal string "#N/A", while an Excel error is
    # represented by cell type ``e``.
    if data_type == "e":
        raise FormulaValueError(f"Excel returned an error value at {worksheet}!{coordinate}: {value}")
    if isinstance(value, float) and not math.isfinite(value):
        raise FormulaValueError(f"Excel returned non-finite numeric value at {worksheet}!{coordinate}")
    if value is not None and not isinstance(value, (str, bool, int, float, date, datetime, datetime_time)):
        raise FormulaValueError(f"Unsupported Excel result type at {worksheet}!{coordinate}: {type(value).__name__}")


def _formula_cache_presence(path: Path) -> dict[tuple[str, str], bool]:
    """Return whether each formula cell has a serialized ``<v>`` cache.

    Excel writes an empty-string result as ``<v></v>`` (which openpyxl exposes
    as ``None`` on some Excel versions), whereas an uncalculated formula has no
    ``<v>`` element.  Keeping this distinction is essential for fail-closed
    calculation verification.
    """
    result: dict[tuple[str, str], bool] = {}
    try:
        with zipfile.ZipFile(path) as archive:
            wb_root = ElementTree.fromstring(archive.read("xl/workbook.xml"))
            rel_root = ElementTree.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
            rel_targets = {}
            for rel in rel_root:
                rid = rel.attrib.get("Id") or rel.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
                target = rel.attrib.get("Target")
                if rid and target:
                    rel_targets[rid] = "xl/" + target.lstrip("/") if not target.startswith("/") else target.lstrip("/")
            for sheet in wb_root.iter():
                if sheet.tag.rsplit("}", 1)[-1] != "sheet":
                    continue
                sheet_name = sheet.attrib.get("name")
                rid = sheet.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
                target = rel_targets.get(rid)
                if not sheet_name or not target or target not in archive.namelist():
                    continue
                root = ElementTree.fromstring(archive.read(target))
                for cell in root.iter():
                    if cell.tag.rsplit("}", 1)[-1] != "c" or not cell.attrib.get("r"):
                        continue
                    if any(child.tag.rsplit("}", 1)[-1] == "f" for child in cell):
                        result[(sheet_name, cell.attrib["r"])] = any(
                            child.tag.rsplit("}", 1)[-1] == "v" for child in cell
                        )
    except (OSError, KeyError, zipfile.BadZipFile, ElementTree.ParseError):
        return {}
    return result


def _materialize_from_result(source: Path, calculated: Path, output: Path,
                             formula_keys: set[tuple[str, str]]) -> Path:
    keep_vba = source.suffix.lower() == ".xlsm"
    try:
        original = load_workbook(source, data_only=False, keep_links=True, keep_vba=keep_vba)
        values = load_workbook(calculated, data_only=True, keep_links=False, keep_vba=keep_vba)
    except Exception as exc:
        raise FormulaValueError(f"Cannot read Excel calculation result: {exc}") from exc
    try:
        cache_presence = _formula_cache_presence(calculated)
        for sheet_name, coordinate in formula_keys:
            if sheet_name not in values.sheetnames or sheet_name not in original.sheetnames:
                raise FormulaValueError(f"Calculation result lost sheet {sheet_name!r}")
            source_cell = original[sheet_name][coordinate]
            result_cell = values[sheet_name][coordinate]
            value = result_cell.value
            _validate_result(value, sheet_name, coordinate, result_cell.data_type)
            # A missing <v> cache is different from a legitimate cached empty
            # string.  Excel's data_only reader returns None in both unusual
            # cases; normal formulas returning "" have a string cache and are
            # accepted.  Missing cache is rejected by the post-save formula
            # count/result verification below if Excel did not calculate it.
            if value is None:
                if result_cell.data_type not in ('s', 'str') or not cache_presence.get((sheet_name, coordinate), False):
                    raise FormulaValueError(f"Excel did not provide a cached result at {sheet_name}!{coordinate}")
                # Excel's empty-string cache can be represented by an empty
                # <v> with no Python value.  It is still a string result.
                value = ""
            _set_literal(source_cell, value)

        output.parent.mkdir(parents=True, exist_ok=True)
        temp_output = output.with_name(f".{output.stem}.tmp-{os.getpid()}-{time.time_ns()}{output.suffix}")
        try:
            original.save(temp_output)
            check = load_workbook(temp_output, data_only=False, keep_links=False, keep_vba=keep_vba)
            try:
                leftovers = [(ws.title, c.coordinate) for ws, c in _iter_formula_cells(check)]
            finally:
                check.close()
            if leftovers:
                raise FormulaValueError(f"Formula values remained after materialization: {leftovers[:3]}")
            os.replace(temp_output, output)
        finally:
            if temp_output.exists():
                temp_output.unlink(missing_ok=True)
    finally:
        original.close()
        values.close()
    return output


def materialize_values(source_path: Path | str, output_path: Path | str) -> Path:
    """Calculate every formula in ``source_path`` and emit a literal snapshot.

    The source is never opened by Excel and is never modified.  A workbook with
    no formulas is returned unchanged, as an explicit fast path.
    """
    source = Path(source_path).resolve()
    output = Path(output_path).resolve()
    if not source.is_file():
        raise FormulaValueError(f"Workbook does not exist: {source}")
    if source == output:
        raise FormulaValueError("Output path must differ from source path")
    if source.suffix.lower() not in {".xlsx", ".xlsm"}:
        raise FormulaValueError(f"Unsupported workbook extension: {source.suffix}")
    if output.suffix.lower() != source.suffix.lower():
        raise FormulaValueError("Output extension must match source extension")

    _scan_package(source)
    keep_vba = source.suffix.lower() == ".xlsm"
    try:
        workbook = load_workbook(source, data_only=False, keep_links=True, keep_vba=keep_vba)
    except Exception as exc:
        raise FormulaValueError(f"Cannot open workbook for formula inspection: {exc}") from exc
    try:
        formulas = _scan_formulas(workbook)
        _scan_defined_names(workbook)
    finally:
        workbook.close()
    if not formulas:
        return source

    timeout_setting = os.environ.get("XLMERGE_EXCEL_TIMEOUT_SECONDS", str(DEFAULT_TIMEOUT_SECONDS))
    try:
        timeout_seconds = max(1, int(timeout_setting))
    except ValueError as exc:
        raise FormulaValueError("XLMERGE_EXCEL_TIMEOUT_SECONDS must be an integer") from exc

    temp_dir = Path(tempfile.mkdtemp(prefix="xlmerge-formula-values-"))
    temporary_source = temp_dir / source.name
    temporary_result = temp_dir / ("calculated" + source.suffix.lower())
    try:
        shutil.copy2(source, temporary_source)
        _run_excel_recalculate(temporary_source, temporary_result, timeout_seconds)
        return _materialize_from_result(source, temporary_result, output,
                                        {(sheet, coordinate) for sheet, coordinate, _ in formulas})
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


__all__ = ["FormulaValueError", "materialize_values"]
