"""
Stage 2 — turn profiled sheets into a semantic model.

This is where a pile of tables becomes something Power BI can reason about:
relationships between sheets, a proper Date table, and DAX measures. The
output (`SemanticModel`) is a pure description — no file writing happens here.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

import pandas as pd

from .profiling import ColumnProfile, TableProfile, UnpivotPlan

# ---------------------------------------------------------------------------
# Model description objects
# ---------------------------------------------------------------------------

DATE_TABLE_NAME = "Date"


@dataclass
class Relationship:
    """A one-to-many relationship. `from_` is the many side (the fact)."""

    from_table: str
    from_column: str
    to_table: str
    to_column: str
    confidence: float = 1.0
    reason: str = ""

    @property
    def name(self) -> str:
        return f"{self.from_table}_{self.from_column}_to_{self.to_table}_{self.to_column}"


@dataclass
class Measure:
    """A DAX measure to write into the model."""

    name: str  # display name, e.g. "Total Revenue"
    dax: str
    format_string: str = "#,##0"
    table: str = ""  # table the measure lives on
    description: str = ""
    is_kpi: bool = False  # should it appear in the headline KPI row?
    # What the measure is made of. Kept separately from the DAX so a visual can
    # bind straight to the column instead, which works whether or not Power BI
    # accepted the measure definition.
    source_column: str = ""
    agg: str = ""  # "SUM" | "AVERAGE" | "COUNTROWS"

    @property
    def safe_name(self) -> str:
        return self.name


@dataclass
class ModelTable:
    name: str
    frame: pd.DataFrame
    columns: list[ColumnProfile]
    is_date_table: bool = False
    hidden: bool = False
    source_sheet: str = ""
    notes: list[str] = field(default_factory=list)
    # Set when the table is written into a workbook the user maintains: the
    # layout that goes on the sheet, and how Power Query reshapes it on refresh.
    excel_frame: pd.DataFrame | None = None
    unpivot_plan: UnpivotPlan | None = None

    def role(self, role: str) -> list[ColumnProfile]:
        return [c for c in self.columns if c.role == role]

    @property
    def sheet_frame(self) -> pd.DataFrame:
        return self.excel_frame if self.excel_frame is not None else self.frame


@dataclass
class SemanticModel:
    tables: list[ModelTable]
    relationships: list[Relationship]
    measures: list[Measure]
    primary_table: str = ""
    primary_date_column: tuple[str, str] | None = None  # (table, column)
    dropped_relationships: list[str] = field(default_factory=list)

    def table(self, name: str) -> ModelTable | None:
        return next((t for t in self.tables if t.name == name), None)

    def measures_on(self, table: str) -> list[Measure]:
        return [m for m in self.measures if m.table == table]

    def to_prompt_dict(self) -> dict:
        """Compact description handed to Claude — structure only, no data rows."""
        return {
            "tables": [
                {
                    "table": t.name,
                    "rows": len(t.frame),
                    "is_date_table": t.is_date_table,
                    "columns": [c.to_prompt_dict() for c in t.columns if c.role != "ignored"],
                }
                for t in self.tables
            ],
            "relationships": [
                f"{r.from_table}[{r.from_column}] -> {r.to_table}[{r.to_column}]"
                for r in self.relationships
            ],
            "measures": [
                {"name": m.name, "table": m.table, "dax": m.dax} for m in self.measures
            ],
        }


# ---------------------------------------------------------------------------
# Relationship detection
# ---------------------------------------------------------------------------


def _normalise(name: str) -> str:
    return re.sub(r"[^a-z0-9]", "", str(name).lower())


def _is_unique_key(frame: pd.DataFrame, col: str) -> bool:
    s = frame[col].dropna()
    return len(s) > 0 and s.is_unique


def detect_relationships(tables: list[ModelTable], min_overlap: float = 0.6) -> list[Relationship]:
    """Find one-to-many joins between tables by matching key columns.

    Two conditions must both hold, which is what keeps this from inventing
    relationships: the column names must match after normalisation, AND the
    values on the many side must actually be found on the one side. Name
    matching alone would happily join "Status" in one sheet to "Status" in
    another that uses a completely different vocabulary.
    """
    rels: list[Relationship] = []
    seen: set[tuple[str, str]] = set()

    for dim in tables:
        if dim.is_date_table:
            continue
        for dim_col in dim.columns:
            if dim_col.role not in ("key", "dimension"):
                continue
            if not _is_unique_key(dim.frame, dim_col.name):
                continue

            dim_values = set(dim.frame[dim_col.name].dropna().astype(str))
            if len(dim_values) < 2:
                continue

            for fact in tables:
                if fact.name == dim.name or fact.is_date_table:
                    continue
                for fact_col in fact.columns:
                    if _normalise(fact_col.name) != _normalise(dim_col.name):
                        continue
                    if _is_unique_key(fact.frame, fact_col.name):
                        continue  # both unique => not a one-to-many
                    fact_values = set(fact.frame[fact_col.name].dropna().astype(str))
                    if not fact_values:
                        continue
                    overlap = len(fact_values & dim_values) / len(fact_values)
                    if overlap < min_overlap:
                        continue
                    pair = (f"{fact.name}.{fact_col.name}", f"{dim.name}.{dim_col.name}")
                    if pair in seen:
                        continue
                    seen.add(pair)
                    rels.append(
                        Relationship(
                            from_table=fact.name, from_column=fact_col.name,
                            to_table=dim.name, to_column=dim_col.name,
                            confidence=overlap,
                            reason=(
                                f"'{fact_col.name}' matches by name and "
                                f"{overlap:.0%} of its values exist in {dim.name}."
                            ),
                        )
                    )
    return rels


class _DisjointSet:
    def __init__(self):
        self.parent: dict[str, str] = {}

    def find(self, x: str) -> str:
        self.parent.setdefault(x, x)
        while self.parent[x] != x:
            self.parent[x] = self.parent[self.parent[x]]
            x = self.parent[x]
        return x

    def union(self, a: str, b: str) -> bool:
        ra, rb = self.find(a), self.find(b)
        if ra == rb:
            return False
        self.parent[ra] = rb
        return True


def prune_relationships(
    relationships: list[Relationship],
    tables: list[ModelTable],
    primary_date: tuple[str, str] | None = None,
) -> tuple[list[Relationship], list[str]]:
    """Reduce the detected joins to a set Power BI will actually accept.

    Power BI allows exactly one *active* filter path between any two tables.
    Raw name-matching happily produces several — a fact joined to a dimension
    on both its ID and its name column, or a loop where Fact→Dim1, Fact→Dim2
    and Dim1→Dim2 all exist. Loading that model gives "ambiguous relationship"
    errors, or silently deactivated relationships that make visuals wrong.

    So we keep a spanning forest: sort the candidates by how much we trust
    them, then accept an edge only if it connects two tables that are not
    already linked by some path. That removes duplicates and cycles in one
    pass, and the discarded ones are reported rather than hidden.
    """
    sizes = {t.name: len(t.frame) for t in tables}
    id_like = re.compile(r"(^|[_\s])(id|key|code|no|num)([_\s]|$)", re.IGNORECASE)

    def priority(r: Relationship) -> tuple:
        is_date = r.to_table == DATE_TABLE_NAME
        is_primary_date = primary_date is not None and (r.from_table, r.from_column) == primary_date
        return (
            1 if is_primary_date else 0,        # the chosen trend date wins outright
            1 if id_like.search(r.from_column) else 0,  # surrogate keys beat name matches
            r.confidence,
            sizes.get(r.from_table, 0),          # the bigger table is the fact side
            0 if is_date else 1,
        )

    ordered = sorted(relationships, key=priority, reverse=True)
    ds = _DisjointSet()
    kept: list[Relationship] = []
    dropped: list[str] = []

    for r in ordered:
        if ds.union(r.from_table, r.to_table):
            kept.append(r)
        else:
            dropped.append(
                f"{r.from_table}[{r.from_column}] → {r.to_table}[{r.to_column}] — "
                f"{r.from_table} and {r.to_table} are already connected by another path, "
                f"and Power BI allows only one."
            )
    return kept, dropped


# ---------------------------------------------------------------------------
# Date table
# ---------------------------------------------------------------------------


def choose_primary_date(tables: list[ModelTable]) -> tuple[str, str] | None:
    """Pick the date column a trend chart should run on.

    Preference: a column named like a transaction date, on the biggest table.
    """
    best = None
    best_score = -1.0
    preferred = re.compile(
        r"(order|invoice|transaction|posting|document|entry|event|sale|"
        r"created|start|booking|receipt|dispatch|delivery)", re.IGNORECASE
    )
    for t in tables:
        if t.is_date_table:
            continue
        for c in t.role("date"):
            score = len(t.frame) / 1000.0
            if preferred.search(c.name):
                score += 50
            if re.search(r"date", c.name, re.IGNORECASE):
                score += 10
            score -= c.null_pct / 10.0
            if score > best_score:
                best_score, best = score, (t.name, c.name)
    return best


def build_date_table(tables: list[ModelTable]) -> ModelTable | None:
    """Build a contiguous calendar table spanning every date in the model."""
    mins, maxs = [], []
    for t in tables:
        for c in t.role("date"):
            if c.min_value is not None and not pd.isna(c.min_value):
                mins.append(pd.Timestamp(c.min_value))
            if c.max_value is not None and not pd.isna(c.max_value):
                maxs.append(pd.Timestamp(c.max_value))
    if not mins or not maxs:
        return None

    start = min(mins).normalize().replace(month=1, day=1)
    end = max(maxs).normalize().replace(month=12, day=31)
    # Guard against a stray 1900 or 2999 date blowing the table up.
    if (end - start).days > 365 * 40:
        start = max(mins).normalize().replace(month=1, day=1) - pd.DateOffset(years=5)
        end = max(maxs).normalize().replace(month=12, day=31)

    idx = pd.date_range(start, end, freq="D")
    df = pd.DataFrame({"Date": idx})
    df["Year"] = df["Date"].dt.year
    df["Quarter"] = "Q" + df["Date"].dt.quarter.astype(str)
    df["Month Number"] = df["Date"].dt.month
    df["Month"] = df["Date"].dt.strftime("%b")
    df["Month Year"] = df["Date"].dt.strftime("%b %Y")
    df["Year Month Sort"] = df["Date"].dt.year * 100 + df["Date"].dt.month
    df["Day"] = df["Date"].dt.day

    cols = [
        ColumnProfile("Date", "date", "date", len(df), 0, len(df),
                      min_value=start, max_value=end, reason="Calendar date."),
        ColumnProfile("Year", "dimension", "number", df["Year"].nunique(), 0, len(df),
                      reason="Calendar year."),
        ColumnProfile("Quarter", "dimension", "text", 4, 0, len(df), reason="Calendar quarter."),
        ColumnProfile("Month Number", "key", "number", 12, 0, len(df), reason="Sort order for Month."),
        ColumnProfile("Month", "dimension", "text", 12, 0, len(df), reason="Month short name."),
        ColumnProfile("Month Year", "dimension", "text", df["Month Year"].nunique(), 0, len(df),
                      reason="Month and year label."),
        ColumnProfile("Year Month Sort", "key", "number", df["Year Month Sort"].nunique(), 0, len(df),
                      reason="Sort order for Month Year."),
        ColumnProfile("Day", "dimension", "number", 31, 0, len(df), reason="Day of month."),
    ]
    return ModelTable(name=DATE_TABLE_NAME, frame=df, columns=cols, is_date_table=True)


# ---------------------------------------------------------------------------
# Measures
# ---------------------------------------------------------------------------

_PERCENT_HINT = re.compile(r"(pct|percent|%|rate|ratio|margin|utilisation|utilization|yield|efficiency)", re.I)
_CURRENCY_HINT = re.compile(r"(amount|amt|revenue|sales|cost|price|value|spend|budget|salary|"
                            r"wage|expense|income|turnover|profit|margin|freight|tax|discount)", re.I)


def _format_string_for(col: ColumnProfile) -> str:
    if _PERCENT_HINT.search(col.name):
        return "0.0%" if (col.max_value is not None and col.max_value <= 1.5) else '#,##0.0"%"'
    if _CURRENCY_HINT.search(col.name):
        return "#,##0.00"
    if col.min_value is not None and col.max_value is not None:
        try:
            if float(col.min_value).is_integer() and float(col.max_value).is_integer():
                return "#,##0"
        except (TypeError, ValueError):
            pass
    return "#,##0.00"


_ABBREVIATIONS = {
    "inr": "INR", "usd": "USD", "eur": "EUR", "gbp": "GBP", "qty": "Quantity",
    "amt": "Amount", "hrs": "Hours", "hr": "Hours", "pct": "%", "no": "No.",
    "id": "ID", "sla": "SLA", "uom": "UOM", "ytd": "YTD", "mtd": "MTD",
    "kg": "kg", "mt": "MT", "poh": "POH",
}


def prettify(name: str) -> str:
    """Turn a database-style column name into something readable on a slide.

    'Activity_Cost_INR' -> 'Activity Cost INR'; 'total_qty' -> 'Total Quantity'.
    """
    s = re.sub(r"[_\-]+", " ", str(name)).strip()
    # Split camelCase, but leave runs of capitals (INR, SLA) alone.
    s = re.sub(r"(?<=[a-z0-9])(?=[A-Z])", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    words = []
    for w in s.split(" "):
        low = w.lower()
        if low in _ABBREVIATIONS:
            words.append(_ABBREVIATIONS[low])
        elif w.isupper() and len(w) > 1:
            words.append(w)  # already an acronym
        else:
            words.append(w[:1].upper() + w[1:] if w else w)
    return " ".join(words)


def _measure_display_name(col_name: str) -> str:
    """Sum of 'Order Amount' reads better as 'Total Order Amount'."""
    n = prettify(col_name)
    if re.match(r"^(total|sum|count|no\.? of|number of|avg|average)\b", n, re.IGNORECASE):
        return n
    if _PERCENT_HINT.search(n):
        return f"Avg {n}"
    return f"Total {n}"


def _dax_quote(table: str, column: str) -> str:
    """DAX column reference — quote the table name if it needs it."""
    if re.match(r"^[A-Za-z_][A-Za-z0-9_]*$", table):
        return f"{table}[{column}]"
    return f"'{table}'[{column}]"


def _count_column(table: ModelTable) -> str:
    """Pick the column to count rows by.

    A row count expressed as "count the non-blank values of column X" is only
    equal to the row count when X is never blank. So prefer a key — which is
    unique and complete by definition — and otherwise take whichever column has
    the fewest blanks, rather than whatever happened to be first.
    """
    candidates = [c for c in table.columns if c.role != "ignored"]
    if not candidates:
        return ""
    keys = [c for c in candidates if c.role == "key" and c.n_null == 0]
    if keys:
        return keys[0].name
    return min(candidates, key=lambda c: (c.n_null, candidates.index(c))).name


def default_measures(tables: list[ModelTable], max_per_table: int = 12) -> list[Measure]:
    """One aggregation per numeric column, plus a row count per table."""
    measures: list[Measure] = []
    used_names: set[str] = set()

    for t in tables:
        if t.is_date_table:
            continue
        numeric = t.role("measure")[:max_per_table]

        for col in numeric:
            agg = "AVERAGE" if _PERCENT_HINT.search(col.name) else "SUM"
            display = _measure_display_name(col.name)
            base = display
            i = 2
            while display in used_names:
                display = f"{base} ({t.name})" if i == 2 else f"{base} {i}"
                i += 1
            used_names.add(display)
            measures.append(
                Measure(
                    name=display,
                    dax=f"{agg}({_dax_quote(t.name, col.name)})",
                    format_string=_format_string_for(col),
                    table=t.name,
                    description=f"{agg.title()} of the {col.name} column.",
                    source_column=col.name,
                    agg=agg,
                )
            )

        count_name = f"{prettify(t.name)} Count"
        if count_name not in used_names:
            used_names.add(count_name)
            measures.append(
                Measure(
                    name=count_name,
                    dax=f"COUNTROWS('{t.name}')",
                    format_string="#,##0",
                    table=t.name,
                    description=f"Number of rows in {t.name}.",
                    source_column=_count_column(t),
                    agg="COUNTROWS",
                )
            )
    return measures


# ---------------------------------------------------------------------------
# Assembly
# ---------------------------------------------------------------------------


def build_model(profiles: list[TableProfile], add_date_table: bool = True) -> SemanticModel:
    tables = [
        ModelTable(
            name=p.name, frame=p.frame, columns=p.columns,
            source_sheet=p.source_sheet, notes=list(p.notes),
            excel_frame=p.excel_frame, unpivot_plan=p.unpivot_plan,
        )
        for p in profiles
    ]

    relationships = detect_relationships(tables)
    primary_date = choose_primary_date(tables)

    if add_date_table:
        date_table = build_date_table(tables)
        if date_table is not None:
            tables.append(date_table)
            # Relate every date column on every table to the calendar. Only the
            # first per table can be active — Power BI allows one active path.
            active_used: set[str] = set()
            for t in tables:
                if t.is_date_table:
                    continue
                for c in t.role("date"):
                    is_primary = primary_date == (t.name, c.name)
                    if t.name in active_used and not is_primary:
                        continue  # skip extra date columns rather than emit inactive clutter
                    active_used.add(t.name)
                    relationships.append(
                        Relationship(
                            from_table=t.name, from_column=c.name,
                            to_table=DATE_TABLE_NAME, to_column="Date",
                            reason="Date column joined to the generated calendar table.",
                        )
                    )

    relationships, dropped = prune_relationships(relationships, tables, primary_date)
    if dropped:
        target = next((t for t in tables if not t.is_date_table), None)
        if target is not None:
            target.notes.append(
                f"Found {len(dropped)} extra table link(s) that would have made the model "
                f"ambiguous, and kept the strongest one in each case."
            )

    measures = default_measures(tables)

    # The primary table is the one with the most rows that carries measures,
    # because that is what a dashboard should headline.
    candidates = [t for t in tables if not t.is_date_table and t.role("measure")]
    if not candidates:
        candidates = [t for t in tables if not t.is_date_table]
    primary = max(candidates, key=lambda t: len(t.frame)).name if candidates else ""

    return SemanticModel(
        tables=tables,
        relationships=relationships,
        measures=measures,
        primary_table=primary,
        primary_date_column=primary_date,
        dropped_relationships=dropped,
    )
