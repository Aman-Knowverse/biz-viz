"""
Stage 3a — write the semantic model as TMDL.

TMDL is the text format Power BI Desktop uses for the model behind a report:
tables, columns, relationships, DAX measures and the Power Query (M) that loads
the data. It is tab-indented and fussy about a handful of things, all of which
are handled in one place here.

Data loading strategy
---------------------
By default the cleaned data is embedded *inside* the model as a compressed
inline table — exactly the mechanism Power BI's own "Enter Data" uses. That
makes the generated project self-contained: it can be zipped, emailed, and
opened on another machine with no broken file paths. Above a size threshold
that stops being sensible, we fall back to shipping CSVs alongside and a
`DataFolder` parameter the user points at once.
"""

from __future__ import annotations

import base64
import json
import re
import uuid
import zlib

import pandas as pd

from .modeling import Measure, ModelTable, Relationship, SemanticModel

TAB = "\t"

# Embedding the data inside the model is much better UX than shipping CSVs with
# a hard-coded path, so the limit is set generously. It exists only to stop a
# genuinely large extract from producing a model file too big to open.
INLINE_CELL_LIMIT = 2_000_000
INLINE_BYTE_LIMIT = 40 * 1024 * 1024  # rough ceiling on the encoded payload


# ---------------------------------------------------------------------------
# Name quoting
# ---------------------------------------------------------------------------

_SAFE_IDENT = re.compile(r"^[A-Za-z_][A-Za-z0-9_]*$")


def q(name: str) -> str:
    """Quote a TMDL identifier if it contains anything but word characters."""
    s = str(name)
    if _SAFE_IDENT.match(s):
        return s
    return "'" + s.replace("'", "''") + "'"


def m_escape(s: str) -> str:
    """Escape a string for a Power Query literal."""
    return str(s).replace('"', '""')


# ---------------------------------------------------------------------------
# Type mapping
# ---------------------------------------------------------------------------


def _is_date_only(series: pd.Series) -> bool:
    if pd.api.types.is_datetime64_any_dtype(series):
        s = series.dropna()
    else:
        s = pd.to_datetime(series, errors="coerce", format="mixed", dayfirst=True).dropna()
    if s.empty:
        return True
    return bool(((s.dt.hour == 0) & (s.dt.minute == 0) & (s.dt.second == 0)).all())


def tmdl_type(col_dtype: str, frame: pd.DataFrame, col_name: str) -> tuple[str, str, str]:
    """Return (tmdl dataType, M type expression, summarizeBy)."""
    if col_dtype == "date":
        if _is_date_only(frame[col_name]):
            return "dateTime", "type date", "none"
        return "dateTime", "type datetime", "none"
    if col_dtype == "boolean":
        return "boolean", "type logical", "none"
    if col_dtype == "number":
        numeric = pd.to_numeric(frame[col_name], errors="coerce").dropna()
        if not numeric.empty and bool(numeric.mod(1).eq(0).all()) and numeric.abs().max() < 9e15:
            return "int64", "Int64.Type", "sum"
        return "double", "type number", "sum"
    return "string", "type text", "none"


# ---------------------------------------------------------------------------
# Serialising a DataFrame into an inline M table
# ---------------------------------------------------------------------------


def _cell_to_text(v, dtype: str) -> str | None:
    """Every inline value is stored as text, then typed by M — same as Enter Data."""
    if v is None or (isinstance(v, float) and pd.isna(v)) or v is pd.NA:
        return None
    try:
        if pd.isna(v):
            return None
    except (TypeError, ValueError):
        pass
    if dtype == "date":
        ts = pd.Timestamp(v)
        if ts.hour or ts.minute or ts.second:
            return ts.strftime("%Y-%m-%dT%H:%M:%S")
        return ts.strftime("%Y-%m-%d")
    if dtype == "boolean":
        return "true" if bool(v) else "false"
    if dtype == "number":
        f = float(v)
        # Invariant formatting: '.' decimal separator, no thousands separator,
        # so the M culture argument below parses it deterministically.
        return str(int(f)) if f.is_integer() and abs(f) < 9e15 else repr(f)
    return str(v)


def encode_inline_table(frame: pd.DataFrame, dtypes: dict[str, str]) -> str:
    """Compress the data the way Power Query's Binary.Decompress expects.

    Power Query uses a *raw* deflate stream (RFC 1951, no zlib header), which is
    why this uses wbits=-15 rather than plain zlib.compress.
    """
    rows = []
    for row in frame.itertuples(index=False, name=None):
        rows.append([_cell_to_text(v, dtypes.get(c, "text")) for v, c in zip(row, frame.columns)])
    payload = json.dumps(rows, separators=(",", ":"), ensure_ascii=False).encode("utf-8")
    compressor = zlib.compressobj(9, zlib.DEFLATED, -15)
    raw = compressor.compress(payload) + compressor.flush()
    return base64.b64encode(raw).decode("ascii")


def inline_partition_m(table: ModelTable, dtypes: dict[str, str], m_types: dict[str, str]) -> str:
    """The M expression for a table whose data lives inside the model file."""
    b64 = encode_inline_table(table.frame, dtypes)
    col_decl = ", ".join(f"{q_m(c)} = _t" for c in table.frame.columns)
    transforms = ", ".join(
        f'{{"{m_escape(c)}", {m_types[c]}}}' for c in table.frame.columns
    )
    return (
        "let\n"
        f'    Source = Table.FromRows(Json.Document(Binary.Decompress(Binary.FromText("{b64}", '
        "BinaryEncoding.Base64), Compression.Deflate)), let _t = ((type nullable text) meta "
        f"[Serialized.Text = true]) in type table [{col_decl}]),\n"
        f'    #"Typed" = Table.TransformColumnTypes(Source,{{{transforms}}}, "en-US")\n'
        "in\n"
        '    #"Typed"'
    )


def _m_list(values) -> str:
    return "{" + ", ".join(f'"{m_escape(v)}"' for v in values) + "}"


def excel_partition_m(
    table: ModelTable,
    m_types: dict[str, str],
    workbook_file: str,
    excel_table: str,
) -> str:
    """The M for a table that reads from the workbook the user maintains.

    Written to be *read*: a consultant will open Power Query and look at these
    steps, so each one is named for what it does rather than left as Source,
    Source1, Source2. The reshaping lives here, not in the file, which is what
    lets the user keep filling in the layout they already know.
    """
    steps: list[str] = []
    steps.append(
        f'    Source = Excel.Workbook(File.Contents(DataFolder & "\\{m_escape(workbook_file)}"), null, true),'
    )
    steps.append(
        f'    RawTable = Source{{[Item="{m_escape(excel_table)}", Kind="Table"]}}[Data],'
    )
    # The sheet ships with spare blank rows so that appending never needs a row
    # insert; they have to be dropped here or they become phantom records.
    steps.append(
        "    DropBlankRows = Table.SelectRows(RawTable, each "
        "List.NonNullCount(Record.FieldValues(_)) > 0),"
    )
    last = "DropBlankRows"

    plan = table.unpivot_plan
    sheet_cols = list(table.sheet_frame.columns)

    # A typed "Total" row is the classic way a hand-maintained sheet starts
    # double-counting. Guard the first text column against it.
    guard = next(
        (c for c in sheet_cols
         if table.sheet_frame[c].dtype == object or str(table.sheet_frame[c].dtype) == "string"),
        None,
    )
    if guard is not None:
        steps.append(
            f'    DropTotalRows = Table.SelectRows({last}, each let v = Record.Field(_, '
            f'"{m_escape(guard)}") in not (v is text and Text.StartsWith(Text.Lower(v), "total"))),'
        )
        last = "DropTotalRows"

    if plan is not None:
        steps.append(
            f"    Unpivoted = Table.UnpivotOtherColumns({last}, "
            f'{_m_list(plan.id_columns)}, "{m_escape(plan.period_name)}", '
            f'"{m_escape(plan.value_name)}"),'
        )
        steps.append(
            f'    DropEmptyCells = Table.SelectRows(Unpivoted, each '
            f'Record.Field(_, "{m_escape(plan.value_name)}") <> null),'
        )
        steps.append(
            f'    AddPeriodOrder = Table.AddColumn(DropEmptyCells, "{m_escape(plan.order_name)}", '
            f'each List.PositionOf({_m_list(plan.period_columns)}, '
            f'Record.Field(_, "{m_escape(plan.period_name)}")), Int64.Type),'
        )
        last = "AddPeriodOrder"

    transforms = ", ".join(
        f'{{"{m_escape(c)}", {m_types[c]}}}' for c in table.frame.columns if c in m_types
    )
    steps.append(f'    SetDataTypes = Table.TransformColumnTypes({last}, {{{transforms}}}, "en-US")')

    return "let\n" + "\n".join(steps) + "\nin\n    SetDataTypes"


def power_query_calendar_m(workbook_file: str, excel_table: str, date_column: str) -> str:
    """A calendar built in Power Query, so it grows as new dates are entered.

    A calendar generated once at build time silently stops covering the data the
    first time someone enters a date in a new year — and the trend chart just
    drops those rows. Rebuilding it on every refresh from the actual minimum and
    maximum in the source is the only version that stays correct unattended.
    """
    return (
        "let\n"
        f'    Source = Excel.Workbook(File.Contents(DataFolder & "\\{m_escape(workbook_file)}"), null, true),\n'
        f'    FactTable = Source{{[Item="{m_escape(excel_table)}", Kind="Table"]}}[Data],\n'
        f'    RawDates = Table.Column(FactTable, "{m_escape(date_column)}"),\n'
        "    UsableDates = List.RemoveNulls(List.Transform(RawDates, each try Date.From(_) otherwise null)),\n"
        "    HasDates = List.Count(UsableDates) > 0,\n"
        "    FirstDay = if HasDates then Date.StartOfYear(List.Min(UsableDates)) "
        "else #date(Date.Year(DateTime.LocalNow()), 1, 1),\n"
        "    LastDay = if HasDates then Date.EndOfYear(List.Max(UsableDates)) "
        "else #date(Date.Year(DateTime.LocalNow()), 12, 31),\n"
        "    DayCount = Duration.Days(LastDay - FirstDay) + 1,\n"
        "    DaySeries = List.Dates(FirstDay, DayCount, #duration(1, 0, 0, 0)),\n"
        '    Calendar = Table.FromList(DaySeries, Splitter.SplitByNothing(), {"Date"}),\n'
        '    TypedDate = Table.TransformColumnTypes(Calendar, {{"Date", type date}}),\n'
        '    AddYear = Table.AddColumn(TypedDate, "Year", each Date.Year([Date]), Int64.Type),\n'
        '    AddQuarter = Table.AddColumn(AddYear, "Quarter", each "Q" & '
        "Text.From(Date.QuarterOfYear([Date])), type text),\n"
        '    AddMonthNumber = Table.AddColumn(AddQuarter, "Month Number", each Date.Month([Date]), Int64.Type),\n'
        '    AddMonth = Table.AddColumn(AddMonthNumber, "Month", each '
        'Date.ToText([Date], [Format="MMM", Culture="en-US"]), type text),\n'
        '    AddMonthYear = Table.AddColumn(AddMonth, "Month Year", each '
        'Date.ToText([Date], [Format="MMM yyyy", Culture="en-US"]), type text),\n'
        '    AddSortKey = Table.AddColumn(AddMonthYear, "Year Month Sort", each '
        "Date.Year([Date]) * 100 + Date.Month([Date]), Int64.Type),\n"
        '    AddDay = Table.AddColumn(AddSortKey, "Day", each Date.Day([Date]), Int64.Type)\n'
        "in\n"
        "    AddDay"
    )


def expressions_tmdl(default_folder: str) -> str:
    """The DataFolder parameter.

    Power Query has no relative paths, so the workbook's location has to be
    stored somewhere. As a parameter it shows up in Power BI's Manage Parameters
    dialog, which means a user who extracted the zip somewhere else fixes it by
    typing a folder — not by editing a query.

    This file needs no `ref` in model.tmdl: Microsoft's TMDL deserialiser
    discovers the root-level files (expressions, relationships, functions,
    dataSources) automatically.
    """
    return (
        f'/// Folder containing the workbook this report reads from.\n'
        f'/// If you moved the files, set this to the new folder.\n'
        f'expression DataFolder = "{m_escape(default_folder.rstrip(chr(92) + "/"))}" meta '
        '[IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true]\n'
    )


def csv_partition_m(table: ModelTable, m_types: dict[str, str], data_folder: str) -> str:
    """The M expression for a table loaded from a CSV shipped beside the project.

    The folder path is written as a literal rather than a Power Query parameter:
    a parameter needs a `ref expression` entry in model.tmdl that the TMDL
    folder loader does not reliably pick up, whereas a literal always loads and
    is a single visible edit in Power Query if the user moves the folder.
    """
    transforms = ", ".join(
        f'{{"{m_escape(c)}", {m_types[c]}}}' for c in table.frame.columns
    )
    path = data_folder.rstrip("\\/") + "\\" + f"{table.name}.csv"
    return (
        "let\n"
        f'    Source = Csv.Document(File.Contents("{m_escape(path)}"), '
        "[Delimiter=\",\", Encoding=65001, QuoteStyle=QuoteStyle.Csv]),\n"
        '    #"Promoted" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),\n'
        f'    #"Typed" = Table.TransformColumnTypes(#"Promoted",{{{transforms}}}, "en-US")\n'
        "in\n"
        '    #"Typed"'
    )


def q_m(name: str) -> str:
    """Quote a column name inside an M `type table [...]` declaration."""
    s = str(name)
    if _SAFE_IDENT.match(s):
        return s
    return "#" + json.dumps(s)


# ---------------------------------------------------------------------------
# TMDL file bodies
# ---------------------------------------------------------------------------


def database_tmdl() -> str:
    return (
        f"database {uuid.uuid4()}\n"
        f"{TAB}compatibilityLevel: 1567\n"
        f"{TAB}compatibilityMode: powerBI\n"
    )


def model_tmdl(model: SemanticModel) -> str:
    lines = [
        "model Model",
        f"{TAB}culture: en-US",
        f"{TAB}defaultPowerBIDataSourceVersion: powerBI_V3",
        f"{TAB}sourceQueryCulture: en-US",
        "",
    ]
    for t in model.tables:
        lines.append(f"ref table {q(t.name)}")
    lines.append("")
    lines.append("ref cultureInfo en-US")
    lines.append("")
    return "\n".join(lines)


def culture_tmdl() -> str:
    return "cultureInfo en-US\n"


def table_tmdl(
    table: ModelTable,
    measures: list[Measure],
    inline: bool,
    data_folder: str = "",
    excel_source: tuple[str, str] | None = None,  # (workbook file, excel table name)
    calendar_source: tuple[str, str, str] | None = None,  # (file, table, date column)
) -> str:
    """Render one table: measures first, then columns, then the partition."""
    frame = table.frame
    dtypes: dict[str, str] = {}
    m_types: dict[str, str] = {}
    tmdl_types: dict[str, str] = {}
    summarize: dict[str, str] = {}

    for col in table.columns:
        if col.name not in frame.columns:
            continue
        dt, mt, summ = tmdl_type(col.dtype, frame, col.name)
        dtypes[col.name] = col.dtype
        m_types[col.name] = mt
        tmdl_types[col.name] = dt
        # Keys and IDs must never be silently summed in a visual.
        summarize[col.name] = "none" if col.role in ("key", "dimension", "date") else summ

    lines = [f"table {q(table.name)}"]
    if table.is_date_table:
        lines.append(f"{TAB}dataCategory: Time")
    lines.append("")

    for m in measures:
        dax = m.dax.strip()
        if "\n" in dax:
            body = "\n".join(f"{TAB}{TAB}{TAB}{ln}" for ln in dax.splitlines())
            lines.append(f"{TAB}measure {q(m.name)} = ```")
            lines.append(body)
            lines.append(f"{TAB}{TAB}{TAB}```")
        else:
            lines.append(f"{TAB}measure {q(m.name)} = {dax}")
        lines.append(f"{TAB}{TAB}formatString: {m.format_string}")
        lines.append(f"{TAB}{TAB}lineageTag: {uuid.uuid4()}")
        lines.append("")

    for col in table.columns:
        if col.name not in frame.columns:
            continue
        lines.append(f"{TAB}column {q(col.name)}")
        lines.append(f"{TAB}{TAB}dataType: {tmdl_types[col.name]}")
        lines.append(f"{TAB}{TAB}lineageTag: {uuid.uuid4()}")
        if col.role == "key" and not table.is_date_table:
            lines.append(f"{TAB}{TAB}isHidden")
        if table.is_date_table and col.name in ("Month Number", "Year Month Sort"):
            lines.append(f"{TAB}{TAB}isHidden")
        lines.append(f"{TAB}{TAB}summarizeBy: {summarize[col.name]}")
        lines.append(f"{TAB}{TAB}sourceColumn: {col.name}")
        if table.is_date_table and col.name == "Month":
            lines.append(f"{TAB}{TAB}sortByColumn: {q('Month Number')}")
        if table.is_date_table and col.name == "Month Year":
            lines.append(f"{TAB}{TAB}sortByColumn: {q('Year Month Sort')}")
        if col.sort_by and col.sort_by in frame.columns:
            lines.append(f"{TAB}{TAB}sortByColumn: {q(col.sort_by)}")
        if table.is_date_table and col.name == "Date":
            lines.append(f"{TAB}{TAB}isKey")
        lines.append("")

    if calendar_source is not None:
        expr = power_query_calendar_m(*calendar_source)
    elif excel_source is not None:
        expr = excel_partition_m(table, m_types, *excel_source)
    elif inline:
        expr = inline_partition_m(table, dtypes, m_types)
    else:
        expr = csv_partition_m(table, m_types, data_folder)
    lines.append(f"{TAB}partition {q(table.name)} = m")
    lines.append(f"{TAB}{TAB}mode: import")
    lines.append(f"{TAB}{TAB}source =")
    for ln in expr.splitlines():
        lines.append(f"{TAB}{TAB}{TAB}{ln}" if ln.strip() else "")
    lines.append("")
    return "\n".join(lines)


def _col_ref(table: str, column: str) -> str:
    return f"{q(table)}.{q(column)}"


def relationships_tmdl(relationships: list[Relationship]) -> str:
    out = []
    for r in relationships:
        out.append(f"relationship {uuid.uuid4()}")
        out.append(f"{TAB}fromColumn: {_col_ref(r.from_table, r.from_column)}")
        out.append(f"{TAB}toColumn: {_col_ref(r.to_table, r.to_column)}")
        out.append("")
    return "\n".join(out)


def definition_pbism() -> str:
    return json.dumps(
        {
            "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/semanticModel/definitionProperties/1.0.0/schema.json",
            "version": "4.2",
            "settings": {"qnaEnabled": True},
        },
        indent=2,
    )


def platform_file(item_type: str, display_name: str) -> str:
    return json.dumps(
        {
            "$schema": "https://developer.microsoft.com/json-schemas/fabric/gitIntegration/platformProperties/2.0.0/schema.json",
            "metadata": {"type": item_type, "displayName": display_name},
            "config": {"version": "2.0", "logicalId": str(uuid.uuid4())},
        },
        indent=2,
    )


# ---------------------------------------------------------------------------
# Whole-model rendering
# ---------------------------------------------------------------------------


def should_inline(model: SemanticModel) -> bool:
    """Can the data live inside the model file, or does it need to ship as CSV?

    Estimated rather than measured: actually encoding every table twice just to
    decide would double the build time on exactly the large inputs where that
    hurts most. ~6 bytes per cell after deflate + base64 is conservative for the
    short, repetitive strings that dominate business extracts.
    """
    cells = sum(len(t.frame) * len(t.frame.columns) for t in model.tables if not t.is_date_table)
    return cells <= INLINE_CELL_LIMIT and cells * 6 <= INLINE_BYTE_LIMIT


def pick_calendar_source(
    model: SemanticModel, excel_tables: dict[str, str]
) -> tuple[str, str] | None:
    """Which table and column the Power Query calendar should read its range from.

    Prefers the date the trend chart runs on. Falls back to any date column that
    actually made it onto a sheet — a date that only exists in the tidy frame
    (Period, say) is no use, because the calendar query reads the workbook.
    """
    candidates: list[tuple[str, str]] = []
    if model.primary_date_column:
        candidates.append(model.primary_date_column)
    for t in model.tables:
        if t.is_date_table or t.name not in excel_tables:
            continue
        for c in t.role("date"):
            candidates.append((t.name, c.name))

    for tname, cname in candidates:
        t = model.table(tname)
        if t is None or tname not in excel_tables:
            continue
        if cname in t.sheet_frame.columns:
            return (tname, cname)
    return None


def render_semantic_model(
    model: SemanticModel,
    default_data_folder: str = r"C:\BizViz\Data",
    mode: str = "snapshot",  # "snapshot" (embedded) | "linked" (reads the workbook)
    workbook_file: str = "",
    excel_tables: dict[str, str] | None = None,  # model table name -> Excel table name
) -> dict[str, str]:
    """Return {relative path: file content} for the .SemanticModel folder."""
    inline = should_inline(model)
    excel_tables = excel_tables or {}
    linked = mode == "linked" and bool(workbook_file) and bool(excel_tables)

    files: dict[str, str] = {
        "definition.pbism": definition_pbism(),
        ".platform": platform_file("SemanticModel", "BizViz Model"),
        "definition/database.tmdl": database_tmdl(),
        "definition/model.tmdl": model_tmdl(model),
        "definition/cultures/en-US.tmdl": culture_tmdl(),
    }

    calendar_pair = pick_calendar_source(model, excel_tables) if linked else None

    if linked:
        files["definition/expressions.tmdl"] = expressions_tmdl(default_data_folder)

    for t in model.tables:
        excel_source = None
        calendar_source = None

        if linked:
            if t.is_date_table and calendar_pair is not None:
                cal_table, cal_column = calendar_pair
                calendar_source = (workbook_file, excel_tables[cal_table], cal_column)
            elif t.name in excel_tables:
                excel_source = (workbook_file, excel_tables[t.name])

        # The generated calendar is always small enough to inline, even when the
        # fact tables are not — it would be silly to ship a CSV of dates.
        t_inline = inline or t.is_date_table
        files[f"definition/tables/{_safe_filename(t.name)}.tmdl"] = table_tmdl(
            t, model.measures_on(t.name), inline=t_inline, data_folder=default_data_folder,
            excel_source=excel_source, calendar_source=calendar_source,
        )

    if model.relationships:
        files["definition/relationships.tmdl"] = relationships_tmdl(model.relationships)

    return files


def _safe_filename(name: str) -> str:
    return re.sub(r'[<>:"/\\|?*]', "_", str(name)).strip() or "Table"
