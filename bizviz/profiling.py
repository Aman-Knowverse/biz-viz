"""
Stage 1 — read the user's Excel and work out what is actually in it.

The job of this module is to turn "a spreadsheet a human made" into a set of
clean, tidy pandas DataFrames plus a description of what each column *means*
in business terms (is it a metric, a category, a date, an ID?).

Everything here is deliberately conservative: when a guess is uncertain the
column is left as a plain attribute rather than silently promoted to a metric,
because a wrong metric produces a confident, wrong dashboard.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Any

import pandas as pd

# ---------------------------------------------------------------------------
# Constants used by the heuristics
# ---------------------------------------------------------------------------

# Words that mark a row as a spreadsheet subtotal rather than a data row.
TOTAL_MARKERS = {"total", "grand total", "subtotal", "sum", "sub total", "grand-total"}

# Column-name fragments that mean "this number is an identifier, not a metric".
ID_NAME_PATTERNS = re.compile(
    r"(^|[\s_\-])(id|ids|no|nos|num|number|code|key|ref|reference|sr|srno|s\.no|serial|"
    r"pin|zip|phone|mobile|gst|pan|invoice|po|so|batch|lot|barcode|sku)([\s_\-]|$)",
    re.IGNORECASE,
)

# Column-name fragments that strongly suggest a real business metric.
MEASURE_NAME_PATTERNS = re.compile(
    r"(amount|amt|value|val|revenue|sales|cost|price|qty|quantity|volume|weight|units|"
    r"total|net|gross|margin|profit|loss|spend|budget|actual|target|plan|forecast|"
    r"count|hours|hrs|days|duration|rate|pct|percent|%|score|balance|stock|inventory|"
    r"discount|tax|freight|salary|wage|expense|income|turnover|output|input|yield|"
    r"efficiency|utilisation|utilization|downtime|uptime|defect|scrap|rejection)",
    re.IGNORECASE,
)

# Column names that look like time periods when a sheet is a crosstab.
MONTH_TOKENS = (
    "jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec|"
    "january|february|march|april|june|july|august|september|october|november|december"
)
PERIOD_HEADER = re.compile(
    rf"^\s*((fy)?\s*(19|20)\d{{2}}([\-/ ]?\d{{2,4}})?|q[1-4]([\s\-/]?(fy)?\s*(19|20)?\d{{2}})?|"
    rf"({MONTH_TOKENS})[\s\-/']*((19|20)?\d{{2}})?|"
    rf"(19|20)\d{{2}}[\s\-/]({MONTH_TOKENS}))\s*$",
    re.IGNORECASE,
)


# ---------------------------------------------------------------------------
# Data structures
# ---------------------------------------------------------------------------


@dataclass
class ColumnProfile:
    """What we believe about a single column."""

    name: str
    role: str  # "measure" | "dimension" | "date" | "key" | "ignored"
    dtype: str  # "number" | "text" | "date" | "boolean"
    n_unique: int
    n_null: int
    n_rows: int
    sample_values: list[Any] = field(default_factory=list)
    min_value: Any = None
    max_value: Any = None
    reason: str = ""  # human-readable justification, shown in the UI
    sort_by: str | None = None  # another column that defines this one's display order

    @property
    def null_pct(self) -> float:
        return (self.n_null / self.n_rows * 100) if self.n_rows else 0.0

    @property
    def unique_pct(self) -> float:
        return (self.n_unique / self.n_rows * 100) if self.n_rows else 0.0

    def to_prompt_dict(self) -> dict:
        """The compact form handed to Claude — no raw data rows, just shape."""
        d = {
            "column": self.name,
            "role": self.role,
            "type": self.dtype,
            "distinct_values": self.n_unique,
            "blank_pct": round(self.null_pct, 1),
        }
        if self.dtype == "number" and self.min_value is not None:
            d["min"] = _jsonable(self.min_value)
            d["max"] = _jsonable(self.max_value)
        if self.role in ("dimension", "key") and self.n_unique <= 25:
            d["values"] = [str(v) for v in self.sample_values[:25]]
        elif self.sample_values:
            d["examples"] = [str(v) for v in self.sample_values[:3]]
        return d


@dataclass
class UnpivotPlan:
    """How to turn a crosstab sheet into rows — carried out by Power Query.

    When the tool ships a workbook the user keeps filling in, the reshaping
    cannot happen once at generation time: the user would have to type in long
    format forever. So the workbook keeps the familiar sideways layout and this
    plan is translated into Power Query steps that run on every refresh.
    """

    id_columns: list[str]  # the columns to keep as they are
    period_columns: list[str]  # the columns whose *headers* are really data
    period_name: str = "Period"
    value_name: str = "Value"
    order_name: str = "Period Order"
    headers_have_year: bool = False


@dataclass
class TableProfile:
    """A cleaned sheet plus everything we worked out about it."""

    name: str  # sanitised table name used in the model
    source_sheet: str  # original sheet name in the workbook
    frame: pd.DataFrame  # tidy form — what the model sees
    columns: list[ColumnProfile]
    notes: list[str] = field(default_factory=list)  # what cleaning did
    # The layout written into the workbook the user maintains. Identical to
    # `frame` unless the sheet was a crosstab, in which case it stays wide and
    # `unpivot_plan` says how Power Query should reshape it.
    excel_frame: pd.DataFrame | None = None
    unpivot_plan: UnpivotPlan | None = None

    @property
    def sheet_frame(self) -> pd.DataFrame:
        return self.excel_frame if self.excel_frame is not None else self.frame

    def by_role(self, role: str) -> list[ColumnProfile]:
        return [c for c in self.columns if c.role == role]

    def column(self, name: str) -> ColumnProfile | None:
        return next((c for c in self.columns if c.name == name), None)

    def to_prompt_dict(self) -> dict:
        return {
            "table": self.name,
            "rows": len(self.frame),
            "columns": [c.to_prompt_dict() for c in self.columns if c.role != "ignored"],
        }


def _jsonable(v: Any) -> Any:
    """Make numpy/pandas scalars safe for json.dumps."""
    if pd.isna(v):
        return None
    if hasattr(v, "item"):
        try:
            return v.item()
        except Exception:
            pass
    if isinstance(v, (pd.Timestamp,)):
        return v.isoformat()
    return v


# ---------------------------------------------------------------------------
# Sheet cleaning
# ---------------------------------------------------------------------------


def _find_header_row(raw: pd.DataFrame, max_scan: int = 12) -> int:
    """Find the row that is most likely the real header.

    Human spreadsheets often start with a title, a blank row, a "Prepared by"
    line and only then the actual column names. We score each of the first few
    rows on how header-like it is: mostly non-blank, mostly text, all distinct.
    """
    best_row, best_score = 0, -1.0
    scan = min(max_scan, len(raw))
    for i in range(scan):
        row = raw.iloc[i]
        non_null = row.notna().sum()
        if non_null < 2:
            continue
        values = [str(v).strip() for v in row.dropna()]
        distinct = len(set(values)) / len(values) if values else 0
        texty = sum(1 for v in row.dropna() if not isinstance(v, (int, float))) / non_null
        fill = non_null / len(row)
        # A header row is wide, distinct, and made of text labels.
        score = fill * 2 + distinct * 1.5 + texty
        if score > best_score:
            best_row, best_score = i, score
    return best_row


def _dedupe_names(names: list[str]) -> list[str]:
    """Excel allows duplicate headers; the model does not."""
    seen: dict[str, int] = {}
    out = []
    for n in names:
        if n in seen:
            seen[n] += 1
            out.append(f"{n}_{seen[n]}")
        else:
            seen[n] = 0
            out.append(n)
    return out


def _clean_column_name(name: Any, index: int) -> str:
    """Turn a raw header cell into a usable column name."""
    s = "" if name is None or (isinstance(name, float) and pd.isna(name)) else str(name)
    s = re.sub(r"\s+", " ", s).strip()
    # Strip characters that break TMDL/DAX identifiers even when quoted.
    s = s.replace('"', "").replace("[", "(").replace("]", ")")
    if not s or s.lower().startswith("unnamed"):
        s = f"Column {index + 1}"
    return s[:100]


def clean_sheet(raw: pd.DataFrame, sheet_name: str) -> tuple[pd.DataFrame, list[str]]:
    """Turn one raw sheet (read with header=None) into a tidy DataFrame."""
    notes: list[str] = []

    # Drop entirely blank rows and columns before looking for the header.
    raw = raw.dropna(how="all").dropna(axis=1, how="all")
    if raw.empty:
        return pd.DataFrame(), ["Sheet is empty."]

    header_idx = _find_header_row(raw)
    if header_idx > 0:
        notes.append(
            f"Treated row {header_idx + 1} as the header — the {header_idx} row(s) above it "
            f"looked like a title block and were skipped."
        )

    header = raw.iloc[header_idx]
    body = raw.iloc[header_idx + 1 :].copy()
    body.columns = _dedupe_names(
        [_clean_column_name(v, i) for i, v in enumerate(header.tolist())]
    )
    body = body.reset_index(drop=True)

    # Drop blank rows/cols again now that the header is consumed.
    before_rows = len(body)
    body = body.dropna(how="all")
    if len(body) < before_rows:
        notes.append(f"Removed {before_rows - len(body)} blank row(s).")

    # Remove spreadsheet subtotal rows — they double-count in any aggregation.
    if len(body.columns):
        first_col = body[body.columns[0]].astype(str).str.strip().str.lower()
        total_mask = first_col.isin(TOTAL_MARKERS) | first_col.str.startswith("grand total")
        if total_mask.any():
            notes.append(
                f"Removed {int(total_mask.sum())} subtotal/total row(s) — Power BI will "
                f"recalculate totals itself, so keeping them would double-count."
            )
            body = body[~total_mask]

    # Forward-fill merged label cells.
    #
    # The signature of a merged grouping column is specific, and worth matching
    # exactly rather than approximately: every value sits in ONE contiguous
    # block, the blocks tile the column from the first row down, and there are
    # few distinct values.Test that precisely and "Region" (North × 5, South × 5…)
    # fills correctly, while a sparse optional column like "Remarks" — scattered
    # one-off values, many distinct — is left alone. Filling that one would not
    # be tidying, it would be inventing data and repeating it down the sheet.
    for col in body.columns:
        s = body[col]
        if s.dtype != object and str(s.dtype) != "string":
            continue
        blank_share = s.isna().mean()
        if not (0.25 <= blank_share < 0.95):
            continue
        filled = s.ffill()
        if filled.isna().any():
            continue  # starts with blanks, so the blocks do not tile — not merged
        distinct = filled.nunique(dropna=True)
        runs = int((filled != filled.shift()).sum())
        if distinct < 2 or distinct >= len(body):
            continue
        if runs != distinct:
            continue  # values recur out of order — scattered data, not merged blocks
        if distinct > max(2, len(body) * 0.5):
            continue
        body[col] = filled
        notes.append(
            f"Filled down merged label cells in '{col}' — {distinct} value(s), each "
            f"covering one block of rows."
        )

    body = body.reset_index(drop=True)
    body = body.convert_dtypes()
    body, numeric_notes = coerce_numeric_text(body)
    notes.extend(numeric_notes)
    return body, notes


# ---------------------------------------------------------------------------
# Numbers that arrived as text
# ---------------------------------------------------------------------------

CURRENCY_CHARS = "₹$€£¥₩﷼"
_NUM_JUNK = re.compile(rf"[{CURRENCY_CHARS},\s '`]")
_PAREN_NEG = re.compile(r"^\((.*)\)$")


def _clean_number_text(v) -> str | None:
    """Strip the decoration accounting exports put around numbers.

    Handles currency symbols, thousands separators (including the Indian
    lakh/crore grouping, which plain comma removal covers), non-breaking
    spaces, and the accounting convention of parenthesising negatives.
    """
    if v is None:
        return None
    s = str(v).strip()
    if not s or s in {"-", "–", "—", "NA", "N/A", "na", "nil", "Nil"}:
        return None
    neg = False
    m = _PAREN_NEG.match(s)
    if m:
        neg, s = True, m.group(1)
    if s.endswith("-"):  # trailing-minus convention from some SAP exports
        neg, s = True, s[:-1]
    s = _NUM_JUNK.sub("", s)
    if s.startswith("-"):
        neg, s = True, s[1:]
    if not s or not re.fullmatch(r"\d*\.?\d+", s):
        return None
    return ("-" + s) if neg else s


def coerce_numeric_text(df: pd.DataFrame) -> tuple[pd.DataFrame, list[str]]:
    """Convert text columns that are really numbers wearing formatting.

    Only converts when nearly every non-blank value cleans up to a number AND
    the column actually looks decorated — a column of plain digits is left
    alone, because those are usually codes ('001234') where stripping the type
    would destroy the leading zeros.
    """
    notes: list[str] = []
    for col in list(df.columns):
        s = df[col]
        if pd.api.types.is_numeric_dtype(s) or pd.api.types.is_datetime64_any_dtype(s):
            continue
        non_null = s.dropna()
        if non_null.empty:
            continue

        raw = non_null.astype(str)
        # Does it look decorated? Currency, separators, percent, parenthesised or
        # trailing-minus negatives. The bar is low on purpose — a single decorated
        # value is enough of a signal, because a column of *undecorated* digits is
        # usually a code ('001234') where converting would eat the leading zeros.
        decorated = raw.str.contains(rf"[{CURRENCY_CHARS}%,()]|-$", regex=True).mean()
        if decorated < 0.1:
            continue

        is_percent = raw.str.strip().str.endswith("%").mean() >= 0.9
        stripped = raw.str.replace("%", "", regex=False) if is_percent else raw
        cleaned = stripped.map(_clean_number_text)
        if cleaned.isna().mean() > 0.1:
            continue

        converted = pd.to_numeric(cleaned, errors="coerce")
        if is_percent:
            converted = converted / 100.0
            notes.append(f"Converted '{col}' from percentage text to a true ratio (12.5% → 0.125).")
        else:
            notes.append(f"Converted '{col}' from formatted text to numbers.")

        out = pd.Series(pd.NA, index=df.index, dtype="Float64")
        out.loc[converted.index] = converted.astype("Float64")
        df[col] = out
    return df, notes


# ---------------------------------------------------------------------------
# Crosstab (wide) detection and unpivoting
# ---------------------------------------------------------------------------


def detect_wide_columns(df: pd.DataFrame) -> list[str]:
    """Return the columns whose *names* look like time periods.

    A sheet with Jan/Feb/Mar or 2023/2024/2025 as column headers is a crosstab:
    the header carries data. Power BI needs that unpivoted into rows.
    """
    period_cols = [c for c in df.columns if PERIOD_HEADER.match(str(c))]
    # Need at least three to be confident it is a period axis, not a stray column.
    if len(period_cols) < 3:
        return []
    # And they should be mostly numeric — a period column holds values.
    numeric_enough = [
        c
        for c in period_cols
        if pd.to_numeric(df[c], errors="coerce").notna().mean() >= 0.5
    ]
    return numeric_enough if len(numeric_enough) >= 3 else []


def unpivot(
    df: pd.DataFrame, period_cols: list[str], value_name: str = "Value"
) -> tuple[pd.DataFrame, dict[str, str]]:
    """Melt period columns into Period/Value rows.

    Returns the melted frame plus the roles that must be forced on the new
    columns. 'Period' is deliberately kept as text when the original headers
    carry no year: parsing a bare "Jan" into a date invents a year that isn't
    in the data, and every trend chart built on it would then be quietly wrong.
    A companion 'Period Order' column preserves the original column order so
    the axis still reads Jan → Dec rather than alphabetically.
    """
    id_cols = [c for c in df.columns if c not in period_cols]
    out = df.melt(
        id_vars=id_cols, value_vars=period_cols, var_name="Period", value_name=value_name
    )
    out = out[out[value_name].notna()]

    order = {c: i for i, c in enumerate(period_cols)}
    out["Period Order"] = out["Period"].map(order)

    has_year = all(re.search(r"(19|20)\d{2}", str(c)) for c in period_cols)
    forced = {"Period Order": "key"}
    if not has_year:
        forced["Period"] = "dimension"
    return out.reset_index(drop=True), forced


# ---------------------------------------------------------------------------
# Column role inference
# ---------------------------------------------------------------------------


def _try_datetime(s: pd.Series) -> pd.Series | None:
    """Parse a column as dates only if nearly all non-blank values parse."""
    non_null = s.dropna()
    if non_null.empty:
        return None
    if pd.api.types.is_datetime64_any_dtype(s):
        return pd.to_datetime(s, errors="coerce")
    # Pure numbers are years/amounts, not dates — never date-parse them.
    if pd.api.types.is_numeric_dtype(non_null):
        return None
    sample = non_null.astype(str).head(200)
    parsed = pd.to_datetime(sample, errors="coerce", format="mixed", dayfirst=True)
    if parsed.notna().mean() >= 0.9:
        return pd.to_datetime(s, errors="coerce", format="mixed", dayfirst=True)
    return None


def profile_column(df: pd.DataFrame, col: str) -> ColumnProfile:
    s = df[col]
    n_rows = len(s)
    n_null = int(s.isna().sum())
    n_unique = int(s.nunique(dropna=True))
    samples = s.dropna().unique()[:25].tolist()

    dt = _try_datetime(s)
    if dt is not None and dt.notna().sum() > 0:
        return ColumnProfile(
            name=col, role="date", dtype="date", n_unique=n_unique, n_null=n_null,
            n_rows=n_rows, sample_values=samples[:3],
            min_value=dt.min(), max_value=dt.max(),
            reason="Values parse as dates.",
        )

    if pd.api.types.is_bool_dtype(s):
        return ColumnProfile(
            name=col, role="dimension", dtype="boolean", n_unique=n_unique, n_null=n_null,
            n_rows=n_rows, sample_values=samples, reason="True/False flag.",
        )

    numeric = pd.to_numeric(s, errors="coerce")
    is_numeric = numeric.notna().sum() >= max(1, (n_rows - n_null) * 0.9) and (n_rows - n_null) > 0

    if is_numeric:
        looks_like_id = bool(ID_NAME_PATTERNS.search(col))
        looks_like_measure = bool(MEASURE_NAME_PATTERNS.search(col))
        nearly_unique = n_unique > 0 and (n_unique / max(1, n_rows - n_null)) > 0.9
        all_integers = bool(numeric.dropna().mod(1).eq(0).all())

        if looks_like_id and not looks_like_measure:
            role, reason = "key", "Numeric, but the name reads as an identifier."
        elif nearly_unique and all_integers and not looks_like_measure:
            role, reason = "key", "Numeric and almost entirely unique — looks like a row ID."
        elif n_unique <= 12 and all_integers and not looks_like_measure:
            role, reason = "dimension", f"Only {n_unique} distinct whole numbers — reads as a category."
        else:
            role = "measure"
            reason = "Name matches a business metric." if looks_like_measure else "Continuous numeric values."
        return ColumnProfile(
            name=col, role=role, dtype="number", n_unique=n_unique, n_null=n_null,
            n_rows=n_rows, sample_values=samples[:5],
            min_value=numeric.min(), max_value=numeric.max(), reason=reason,
        )

    # Text from here on.
    non_null = n_rows - n_null
    if non_null == 0:
        return ColumnProfile(
            name=col, role="ignored", dtype="text", n_unique=0, n_null=n_null,
            n_rows=n_rows, reason="Column is entirely blank.",
        )

    uniq_ratio = n_unique / non_null
    if uniq_ratio > 0.95 and n_unique > 50:
        role = "key"
        reason = "Text and almost entirely unique — an ID or free-text field, not a category."
    elif n_unique > 1000:
        role = "key"
        reason = f"{n_unique:,} distinct text values — too many to slice by."
    else:
        role = "dimension"
        reason = f"{n_unique} distinct text values — usable as a category."

    return ColumnProfile(
        name=col, role=role, dtype="text", n_unique=n_unique, n_null=n_null,
        n_rows=n_rows, sample_values=samples, reason=reason,
    )


# ---------------------------------------------------------------------------
# Table names
# ---------------------------------------------------------------------------


def sanitise_table_name(name: str, taken: set[str] | None = None) -> str:
    """Make a sheet name safe to use as a model table name."""
    s = re.sub(r"[^A-Za-z0-9 _]+", " ", str(name))
    s = re.sub(r"\s+", " ", s).strip()
    if not s:
        s = "Table"
    if s[0].isdigit():
        s = "T " + s
    s = s[:80]
    if taken is not None:
        base, i = s, 2
        while s in taken:
            s = f"{base} {i}"
            i += 1
        taken.add(s)
    return s


# ---------------------------------------------------------------------------
# Top-level entry point
# ---------------------------------------------------------------------------


def profile_workbook(
    source: Any,
    min_rows: int = 2,
    auto_unpivot: bool = True,
) -> list[TableProfile]:
    """Read every sheet of an Excel/CSV source and profile it.

    `source` may be a path or a file-like object (Streamlit upload).
    """
    name = getattr(source, "name", str(source))
    if str(name).lower().endswith((".csv", ".txt", ".tsv")):
        sep = "\t" if str(name).lower().endswith(".tsv") else None
        raw_sheets = {
            "Data": pd.read_csv(source, header=None, sep=sep, engine="python", dtype=object)
        }
    else:
        raw_sheets = pd.read_excel(source, sheet_name=None, header=None, dtype=object)

    profiles: list[TableProfile] = []
    taken: set[str] = set()

    for sheet_name, raw in raw_sheets.items():
        df, notes = clean_sheet(raw, sheet_name)
        if df.empty or len(df) < min_rows or len(df.columns) < 2:
            continue

        forced: dict[str, str] = {}
        wide_df: pd.DataFrame | None = None
        plan: UnpivotPlan | None = None
        if auto_unpivot:
            period_cols = detect_wide_columns(df)
            if period_cols:
                wide_df = df.copy()  # keep the layout the user recognises
                plan = UnpivotPlan(
                    id_columns=[c for c in df.columns if c not in period_cols],
                    period_columns=list(period_cols),
                    headers_have_year=all(
                        re.search(r"(19|20)\d{2}", str(c)) for c in period_cols
                    ),
                )
                df, forced = unpivot(df, period_cols)
                notes.append(
                    f"Unpivoted {len(period_cols)} period columns "
                    f"({', '.join(map(str, period_cols[:4]))}{'…' if len(period_cols) > 4 else ''}) "
                    f"into Period/Value rows — Power BI needs one row per period, not one column."
                )
                if "Period" in forced:
                    notes.append(
                        "Kept 'Period' as a text category because the original headers had no "
                        "year in them — turning 'Jan' into a date would have invented one."
                    )

        cols = []
        for c in df.columns:
            prof = profile_column(df, c)
            if c in forced:
                prof.role = forced[c]
                if c == "Period":
                    prof.dtype = "text"
                    prof.sort_by = "Period Order"
                    prof.reason = "Period label from the original column headers."
            cols.append(prof)
        profiles.append(
            TableProfile(
                name=sanitise_table_name(sheet_name, taken),
                source_sheet=str(sheet_name),
                frame=df,
                columns=cols,
                notes=notes,
                excel_frame=wide_df,
                unpivot_plan=plan,
            )
        )

    return profiles
