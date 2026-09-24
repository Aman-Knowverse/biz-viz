"""
Stage 3d — write the master workbook the user keeps filling in.

This file is the thing that outlives the dashboard. It is not a dump of cleaned
data; it is a form. Every decision here is about making the next person's typing
land somewhere Power BI can still read six months from now:

* each sheet is a real Excel **Table**, because that is what makes a new row get
  picked up on refresh without anyone editing a query;
* the table already contains spare blank rows, so appending never needs a row
  insert — which is what would otherwise fight with sheet protection;
* category columns get dropdowns, because one "Nort" among the "North"s quietly
  grows a fifth region on the dashboard;
* headers are locked, because a renamed column breaks the model outright.
"""

from __future__ import annotations

import re
import textwrap
from dataclasses import dataclass

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Protection, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.workbook.defined_name import DefinedName

from .modeling import ModelTable

LISTS_SHEET = "_BizViz_Lists"
README_SHEET = "READ ME FIRST"

# Dropdowns only make sense for a column with a small, closed set of values.
MAX_DROPDOWN_VALUES = 50
# Above this many cells, per-cell unlocking costs more than it is worth — and a
# sheet that big is a system export, not something anyone maintains by hand.
PROTECTION_CELL_LIMIT = 100_000

HEADER_FILL = PatternFill("solid", fgColor="12253F")
HEADER_FONT = Font(color="FFFFFF", bold=True, size=11)
TITLE_FONT = Font(color="12253F", bold=True, size=14)
BODY_FONT = Font(size=11)
THIN = Side(style="thin", color="D8DEE4")


@dataclass
class SheetPlan:
    """What was written for one table, so the caller can describe it."""

    table_name: str
    sheet_name: str
    excel_table_name: str
    columns: list[str]
    n_rows: int
    n_spare_rows: int
    dropdown_columns: list[str]
    protected: bool


# ---------------------------------------------------------------------------
# Naming
# ---------------------------------------------------------------------------

_ILLEGAL_SHEET = re.compile(r"[\[\]:*?/\\]")


def safe_sheet_name(name: str, taken: set[str]) -> str:
    s = _ILLEGAL_SHEET.sub(" ", str(name)).strip()
    s = re.sub(r"\s+", " ", s)[:31] or "Sheet"
    base, i = s, 2
    while s.lower() in {t.lower() for t in taken}:
        suffix = f" {i}"
        s = base[: 31 - len(suffix)] + suffix
        i += 1
    taken.add(s)
    return s


def safe_table_name(name: str, taken: set[str]) -> str:
    """Excel table names allow no spaces and must start with a letter."""
    s = re.sub(r"[^A-Za-z0-9_]", "_", str(name)).strip("_")
    s = re.sub(r"_+", "_", s) or "Table"
    if not s[0].isalpha():
        s = "t_" + s
    s = ("tbl_" + s)[:60]
    base, i = s, 2
    while s.lower() in {t.lower() for t in taken}:
        s = f"{base}_{i}"
        i += 1
    taken.add(s)
    return s


def _defined_name_for(table: str, column: str, taken: set[str]) -> str:
    s = re.sub(r"[^A-Za-z0-9_]", "_", f"vals_{table}_{column}").strip("_")
    s = re.sub(r"_+", "_", s)[:80]
    if not s[0].isalpha():
        s = "v_" + s
    base, i = s, 2
    while s.lower() in {t.lower() for t in taken}:
        s = f"{base}_{i}"
        i += 1
    taken.add(s)
    return s


# ---------------------------------------------------------------------------
# Formatting helpers
# ---------------------------------------------------------------------------


def _column_format(series: pd.Series) -> str | None:
    if pd.api.types.is_datetime64_any_dtype(series):
        s = series.dropna()
        has_time = bool(len(s)) and not (
            (s.dt.hour == 0) & (s.dt.minute == 0) & (s.dt.second == 0)
        ).all()
        return "dd-mmm-yyyy hh:mm" if has_time else "dd-mmm-yyyy"
    if pd.api.types.is_numeric_dtype(series) and not pd.api.types.is_bool_dtype(series):
        s = pd.to_numeric(series, errors="coerce").dropna()
        if len(s) and bool(s.mod(1).eq(0).all()):
            return "#,##0"
        return "#,##0.00"
    return None


def _cell_value(v):
    """openpyxl accepts only plain Python scalars."""
    if v is None or v is pd.NA:
        return None
    try:
        if pd.isna(v):
            return None
    except (TypeError, ValueError):
        pass
    if isinstance(v, pd.Timestamp):
        return v.to_pydatetime()
    if hasattr(v, "item"):
        try:
            return v.item()
        except Exception:
            pass
    return v


def _spare_rows(n_rows: int) -> int:
    """Blank rows left inside the table so appending never inserts a row."""
    return max(50, min(500, int(n_rows * 0.1)))


# ---------------------------------------------------------------------------
# The workbook
# ---------------------------------------------------------------------------


def write_master_workbook(
    tables: list[ModelTable],
    path: str,
    project_name: str = "BizViz",
    guard_rails: bool = True,
) -> list[SheetPlan]:
    """Write the workbook and return a description of each sheet written."""
    wb = Workbook()
    wb.remove(wb.active)  # drop the default sheet; we add our own in order

    readme = wb.create_sheet(README_SHEET)
    lists_ws = wb.create_sheet(LISTS_SHEET)

    sheet_names: set[str] = {README_SHEET, LISTS_SHEET}
    table_names: set[str] = set()
    defined_names: set[str] = set()
    plans: list[SheetPlan] = []
    lists_col = 1

    data_tables = [t for t in tables if not t.is_date_table]

    for t in data_tables:
        frame = t.sheet_frame.reset_index(drop=True)
        if frame.empty or not len(frame.columns):
            continue

        sheet = safe_sheet_name(t.name, sheet_names)
        ws = wb.create_sheet(sheet)
        xl_table = safe_table_name(t.name, table_names)

        n_rows, n_cols = len(frame), len(frame.columns)
        spare = _spare_rows(n_rows)

        # --- header
        for j, col in enumerate(frame.columns, start=1):
            c = ws.cell(row=1, column=j, value=str(col))
            c.fill = HEADER_FILL
            c.font = HEADER_FONT
            c.alignment = Alignment(vertical="center", wrap_text=False)
            c.border = Border(bottom=THIN)

        # --- body
        cell_count = n_rows * n_cols
        unlock = guard_rails and cell_count <= PROTECTION_CELL_LIMIT
        unlocked = Protection(locked=False)

        for j, col in enumerate(frame.columns, start=1):
            fmt = _column_format(frame[col])
            values = frame[col].tolist()
            for i, v in enumerate(values, start=2):
                c = ws.cell(row=i, column=j, value=_cell_value(v))
                if fmt:
                    c.number_format = fmt
                if unlock:
                    c.protection = unlocked
            # spare rows carry the format so new entries look right immediately
            for i in range(n_rows + 2, n_rows + 2 + spare):
                c = ws.cell(row=i, column=j)
                if fmt:
                    c.number_format = fmt
                if unlock:
                    c.protection = unlocked

        # --- the Excel Table itself, stretched over the spare rows
        last_row = 1 + n_rows + spare
        ref = f"A1:{get_column_letter(n_cols)}{last_row}"
        table = Table(displayName=xl_table, ref=ref)
        table.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium2", showRowStripes=True, showColumnStripes=False,
            showFirstColumn=False, showLastColumn=False,
        )
        ws.add_table(table)

        # --- column widths and a frozen header
        for j, col in enumerate(frame.columns, start=1):
            sample = frame[col].dropna().astype(str).head(50)
            longest = max([len(str(col))] + [len(v) for v in sample]) if len(sample) else len(str(col))
            ws.column_dimensions[get_column_letter(j)].width = min(38, max(11, longest + 3))
        ws.freeze_panes = "A2"

        # --- dropdowns on the closed-set category columns
        dropdown_cols: list[str] = []
        if guard_rails:
            for prof in t.columns:
                if prof.role != "dimension" or prof.name not in frame.columns:
                    continue
                if prof.dtype not in ("text", "boolean"):
                    continue
                values = sorted(
                    {str(v) for v in frame[prof.name].dropna().unique() if str(v).strip()}
                )
                if not (2 <= len(values) <= MAX_DROPDOWN_VALUES):
                    continue

                # The allowed values live on a hidden sheet, referenced through a
                # name that grows with the list — so adding a genuinely new
                # category is one row on that sheet, not a rebuild.
                letter = get_column_letter(lists_col)
                lists_ws.cell(row=1, column=lists_col, value=f"{t.name} · {prof.name}")
                for k, v in enumerate(values, start=2):
                    lists_ws.cell(row=k, column=lists_col, value=v)

                dn = _defined_name_for(t.name, prof.name, defined_names)
                formula = (
                    f"OFFSET('{LISTS_SHEET}'!${letter}$2,0,0,"
                    f"MAX(1,COUNTA('{LISTS_SHEET}'!${letter}:${letter})-1),1)"
                )
                try:
                    wb.defined_names[dn] = DefinedName(dn, attr_text=formula)
                except TypeError:  # openpyxl < 3.1
                    wb.defined_names.append(DefinedName(dn, attr_text=formula))

                col_letter = get_column_letter(list(frame.columns).index(prof.name) + 1)
                dv = DataValidation(
                    type="list", formula1=f"={dn}", allow_blank=True,
                    showDropDown=False,  # False here means "do show the in-cell arrow"
                    errorStyle="warning",
                    errorTitle="Not a value we've seen before",
                    error=(
                        "This is not one of the existing values for this column. "
                        "Choose Yes only if it really is a new category — otherwise "
                        "a typo here will show up as an extra item on the dashboard."
                    ),
                    promptTitle=prof.name,
                    prompt="Pick from the list, or type a genuinely new value.",
                )
                ws.add_data_validation(dv)
                dv.add(f"{col_letter}2:{col_letter}{last_row}")
                dropdown_cols.append(prof.name)
                lists_col += 1

        # --- lock the headers, leave the data editable
        if unlock:
            # Deliberately no password: this is a guard against a slip, not a
            # lock. Review > Unprotect Sheet undoes it with no ceremony.
            ws.protection.sheet = True
            ws.protection.insertRows = True
            ws.protection.deleteRows = True
            ws.protection.sort = True
            ws.protection.autoFilter = True
            ws.protection.formatCells = True
            ws.protection.formatColumns = True
            ws.protection.selectLockedCells = True
            ws.protection.selectUnlockedCells = True

        plans.append(
            SheetPlan(
                table_name=t.name, sheet_name=sheet, excel_table_name=xl_table,
                columns=[str(c) for c in frame.columns], n_rows=n_rows,
                n_spare_rows=spare, dropdown_columns=dropdown_cols, protected=unlock,
            )
        )

    _write_readme(readme, project_name, plans, tables)
    lists_ws.sheet_state = "hidden"
    if lists_col == 1:
        lists_ws.cell(row=1, column=1, value="No dropdown lists were needed.")

    wb.save(path)
    return plans


# ---------------------------------------------------------------------------
# The instruction sheet
# ---------------------------------------------------------------------------


def _write_readme(ws, project_name: str, plans: list[SheetPlan], tables: list[ModelTable]) -> None:
    ws.column_dimensions["A"].width = 4
    ws.column_dimensions["B"].width = 96
    ws.sheet_view.showGridLines = False

    row = 2
    WRAP_AT = 92

    def line(text: str = "", *, bold=False, size=11, color="1F2933", gap=0):
        """Write one paragraph, wrapped across rows.

        Deliberately not relying on Excel's wrap_text: a wrapped cell keeps the
        default row height, so anything past the first line is simply invisible
        unless the row height is computed — which openpyxl cannot do. Wrapping
        into real rows is uglier code and a document that always displays.
        """
        nonlocal row
        chunks = textwrap.wrap(text, WRAP_AT, subsequent_indent="   ") if text else [""]
        for chunk in chunks:
            c = ws.cell(row=row, column=2, value=chunk)
            c.font = Font(bold=bold, size=size, color=color)
            c.alignment = Alignment(vertical="top")
            row += 1
        row += gap

    def heading(text: str):
        nonlocal row
        row += 1
        c = ws.cell(row=row, column=2, value=text)
        c.font = Font(bold=True, size=12, color="0E7C7B")
        row += 1

    ws.cell(row=row, column=2, value=f"{project_name} — data file").font = TITLE_FONT
    row += 2
    line(
        "This workbook is the single source your Power BI dashboard reads from. "
        "Keep adding data here and the dashboard stays current.",
        size=11,
    )

    heading("To add new data")
    line("1.  Open the sheet you want to add to.")
    line("2.  Type into the first empty row inside the coloured table. Blank rows are already "
         "there waiting — you do not need to insert anything.")
    line("3.  Save and close this file.")
    line("4.  In Power BI, click Home → Refresh. Your new rows appear.")

    heading("Four things that will break the dashboard")
    line("•  Renaming a column heading. The dashboard looks these up by name.")
    line("•  Inserting or deleting a column in the middle of a table.")
    line("•  Adding a 'Total' row. Power BI calculates its own totals — a typed one gets "
         "counted twice.")
    line("•  Typing outside the coloured table. Anything below or beside it is ignored.")

    heading("About the dropdowns")
    line("Category columns show a list of the values already in use. If you type something new "
         "you get a warning, not a block — say yes if it really is new. To add a value to a "
         "list permanently, unhide the sheet called '_BizViz_Lists' and add it to the bottom "
         "of the matching column.")

    if any(p.protected for p in plans):
        heading("Why some cells will not let you type")
        line("The heading row is locked so it cannot be renamed by accident. Everything else is "
             "editable. If you genuinely need to change a heading, go to Review → Unprotect Sheet "
             "— there is no password — but be aware the dashboard will need repointing afterwards.")

    heading("What is in this workbook")
    for p in plans:
        t = next((x for x in tables if x.name == p.table_name), None)
        extra = ""
        if t is not None and t.unpivot_plan is not None:
            extra = (
                f"  Kept in its original side-by-side layout ("
                f"{len(t.unpivot_plan.period_columns)} period columns) — Power BI reshapes it "
                f"on refresh, so keep filling it the way you always have."
            )
        line(f"•  '{p.sheet_name}'  —  {p.n_rows:,} rows, {len(p.columns)} columns, "
             f"{p.n_spare_rows} blank rows ready for new entries.{extra}")

    heading("If you move this file")
    line("Power BI remembers where this file lives by its full path. If you move or rename it, "
         "open Power BI and go to Home → Transform data → Manage Parameters, and set "
         "'DataFolder' to the new folder. Nothing else needs changing.")
