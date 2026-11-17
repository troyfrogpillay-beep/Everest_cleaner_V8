import io
import re
import difflib
from copy import copy
from dataclasses import dataclass, field
from datetime import datetime
from typing import Tuple, List, Dict, Any, Optional, NamedTuple

import openpyxl
from openpyxl.utils import get_column_letter
import pandas as pd
from email_validator import validate_email, EmailNotValidError
import streamlit as st

st.set_page_config(page_title="Everest Data Cleaner", layout="wide")

# ==============================================================================
# PRE-CLEANING STEPS: CLIENT SECTION EXTRACTION
# ==============================================================================

SUMMARY_SHEET_NAME = "Client Premium Summary"
DETAIL_SHEET_NAME = "Client Detailed Schedule"

SUMMARY_HEADERS = ["Payment Date", "Month End", "Policy Name", "Client Name", "Risk Payment"]
DETAIL_HEADERS = ["Client Name", "Member ID", "Surname", "First Name", "Policy Description", "Risk Premium"]


class ExtractionError(Exception):
    """Raised with a user-readable message when an extraction cannot proceed."""


class CellSnapshot(NamedTuple):
    value: Any
    data_type: str
    number_format: str
    font: Any
    fill: Any
    border: Any
    alignment: Any
    protection: Any


@dataclass
class ExtractedSection:
    section_label: str
    default_sheet_name: str
    source_sheet: str
    header_row: int
    first_row: int
    last_row: int
    first_col: int
    last_col: int
    header: List[CellSnapshot] = field(default_factory=list)
    rows: List[List[CellSnapshot]] = field(default_factory=list)
    source_rows: List[int] = field(default_factory=list)
    col_widths: Dict[int, float] = field(default_factory=dict)

    @property
    def section_range(self) -> str:
        return f"'{self.source_sheet}'!{get_column_letter(self.first_col)}{self.header_row}:{get_column_letter(self.last_col)}{self.last_row}"


def _norm_header(val: Any) -> str:
    if not isinstance(val, str):
        return ""
    return re.sub(r"\s+", " ", val.strip()).casefold()


def _norm_client(val: Any) -> str:
    if val is None:
        return ""
    return str(val).strip().casefold()


def _is_blank(val: Any) -> bool:
    return val is None or (isinstance(val, str) and not val.strip())


def _sheet_grid(ws) -> List[tuple]:
    return list(ws.iter_rows(values_only=True))


def _header_positions(row: tuple) -> Dict[str, List[int]]:
    pos: Dict[str, List[int]] = {}
    for idx, v in enumerate(row, start=1):
        key = _norm_header(v)
        if key:
            pos.setdefault(key, []).append(idx)
    return pos


def _matching_header_rows(grid: List[tuple], required: List[str]) -> List[Tuple[int, Dict[str, List[int]]]]:
    req = [_norm_header(h) for h in required]
    hits = []
    for r_idx, row in enumerate(grid, start=1):
        pos = _header_positions(row)
        if all(h in pos for h in req):
            hits.append((r_idx, pos))
    return hits


def _best_partial(grid: List[tuple], required: List[str]) -> Tuple[int, int, List[str]]:
    req = [(h, _norm_header(h)) for h in required]
    best = (0, 0, [h for h, _ in req])
    for r_idx, row in enumerate(grid, start=1):
        pos = _header_positions(row)
        found = [h for h, n in req if n in pos]
        if len(found) > best[1]:
            best = (r_idx, len(found), [h for h, n in req if n not in pos])
    return best


def _extract_section(wb, client_name: str, required: List[str], other_required: List[str], label: str, default_sheet_name: str, stop_at_first_blank_row: bool) -> ExtractedSection:
    if client_name is None or not str(client_name).strip():
        raise ExtractionError(f"Client name is blank. Enter a client name to extract the {label}.")
    target = _norm_client(client_name)

    found: List[Tuple[Any, List[tuple], int, Dict[str, List[int]]]] = []
    best_overall = (None, 0, 0, list(required))
    for ws in wb.worksheets:
        grid = _sheet_grid(ws)
        for r_idx, pos in _matching_header_rows(grid, required):
            found.append((ws, grid, r_idx, pos))
        r, cnt, missing = _best_partial(grid, required)
        if cnt > best_overall[2]:
            best_overall = (ws.title, r, cnt, missing)

    if not found:
        if best_overall[2] >= 2:
            raise ExtractionError(f"Could not find the {label} header. The closest row is '{best_overall[0]}' row {best_overall[1]}, which is missing: {', '.join(best_overall[3])}. Required headers: {', '.join(required)}.")
        raise ExtractionError(f"Could not find the {label} header. No row contains the required headers: {', '.join(required)}.")

    if len(found) > 1:
        where = ", ".join(f"'{ws.title}' row {r}" for ws, _, r, _ in found)
        raise ExtractionError(f"Ambiguous {label} header: the required headers appear in more than one place ({where}). Cannot tell which one to use.")

    ws, grid, header_row, pos = found[0]
    dupes = [h for h in required if len(pos[_norm_header(h)]) > 1]
    if dupes:
        raise ExtractionError(f"Ambiguous {label} header in '{ws.title}' row {header_row}: duplicate column(s) {', '.join(dupes)}.")

    header_row_vals = grid[header_row - 1]
    used_cols = [i for i, v in enumerate(header_row_vals, start=1) if not _is_blank(v)]
    first_col, last_col = min(used_cols), max(used_cols)
    client_col = pos[_norm_header("Client Name")][0]

    other_headers = [r for r, _ in _matching_header_rows(grid, other_required) if r > header_row]
    limit = (min(other_headers) - 1) if other_headers else len(grid)
    first_row = header_row + 1

    def row_blank(r: int) -> bool:
        return all(_is_blank(v) for v in grid[r - 1][first_col - 1:last_col])

    if stop_at_first_blank_row:
        end = first_row - 1
        for r in range(first_row, limit + 1):
            if row_blank(r):
                break
            end = r
    else:
        end = limit
        while end >= first_row and row_blank(end):
            end -= 1

    section = ExtractedSection(section_label=label, default_sheet_name=default_sheet_name, source_sheet=ws.title, header_row=header_row, first_row=first_row, last_row=max(end, header_row), first_col=first_col, last_col=last_col)

    def snap(r: int, c: int) -> CellSnapshot:
        cell = ws.cell(row=r, column=c)
        return CellSnapshot(cell.value, cell.data_type, cell.number_format, copy(cell.font), copy(cell.fill), copy(cell.border), copy(cell.alignment), copy(cell.protection))

    section.header = [snap(header_row, c) for c in range(first_col, last_col + 1)]
    seen_clients: Dict[str, str] = {}
    for r in range(first_row, end + 1):
        row = grid[r - 1]
        raw = row[client_col - 1] if client_col - 1 < len(row) else None
        key = _norm_client(raw)
        if key:
            seen_clients.setdefault(key, str(raw).strip())
        if key == target:
            section.rows.append([snap(r, c) for c in range(first_col, last_col + 1)])
            section.source