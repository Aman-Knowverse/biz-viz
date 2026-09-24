"""
Stage 3b — decide what the dashboard should show.

Two designers live here:

* `rule_based_design` — deterministic, offline, no API key. Picks the largest
  fact table, its best measures and dimensions, and lays out three pages.
* `claude_design` — asks Claude to read the *shape* of the model (never the
  data itself) and return a structured spec: which measures are the headline
  KPIs, what to call them in business language, and how to break the pages up.

Both return the same `ReportSpec`, so the file writer downstream neither knows
nor cares which one ran. That separation is the point: the model can change the
plan, but it never writes a byte of the output.
"""

from __future__ import annotations

import json
import re
from dataclasses import dataclass, field

from .modeling import DATE_TABLE_NAME, Measure, SemanticModel

CANVAS_W, CANVAS_H = 1280, 720


# ---------------------------------------------------------------------------
# Spec objects — the contract between "what to show" and "how to write it"
# ---------------------------------------------------------------------------


@dataclass
class VisualSpec:
    kind: str  # card | line | bar | column | matrix | table | slicer | title
    x: int
    y: int
    width: int
    height: int
    title: str = ""
    measures: list[str] = field(default_factory=list)  # measure display names
    category: tuple[str, str] | None = None  # (table, column)
    series: tuple[str, str] | None = None
    columns: list[tuple[str, str]] = field(default_factory=list)  # for table visual
    text: str = ""  # for title/textbox


@dataclass
class PageSpec:
    title: str
    visuals: list[VisualSpec] = field(default_factory=list)


@dataclass
class ReportSpec:
    title: str
    pages: list[PageSpec] = field(default_factory=list)
    narrative: str = ""  # optional written summary from Claude
    designed_by: str = "rules"


# ---------------------------------------------------------------------------
# Shared helpers
# ---------------------------------------------------------------------------


def _score_dimension(col, n_rows: int) -> float:
    """How useful is this column as a chart category or slicer?

    The sweet spot is 3–20 distinct values. One value tells you nothing; 500
    values make an unreadable bar chart.
    """
    n = col.n_unique
    if n < 2 or n > 60:
        return -1.0
    if 3 <= n <= 12:
        base = 10.0
    elif n <= 25:
        base = 7.0
    else:
        base = 3.0
    if re.search(r"(region|zone|state|country|city|plant|site|location|branch)", col.name, re.I):
        base += 4
    if re.search(r"(category|type|segment|group|class|department|dept|team|status|stage|priority)",
                 col.name, re.I):
        base += 3
    if re.search(r"(product|customer|vendor|supplier|channel|material|item)", col.name, re.I):
        base += 2
    base -= col.null_pct / 20.0
    return base


def rank_dimensions(model: SemanticModel, table_name: str) -> list[tuple[str, str, float]]:
    """Return (table, column, score) for slicer/category candidates.

    Dimensions on *related* tables count too — that is the whole point of
    having built relationships.
    """
    out: list[tuple[str, str, float]] = []
    tables = {table_name}
    for r in model.relationships:
        if r.from_table == table_name and r.to_table != DATE_TABLE_NAME:
            tables.add(r.to_table)

    for tname in tables:
        t = model.table(tname)
        if t is None:
            continue
        for c in t.columns:
            if c.role != "dimension":
                continue
            s = _score_dimension(c, len(t.frame))
            if s > 0:
                out.append((tname, c.name, s))
    out.sort(key=lambda x: -x[2])
    return out


def rank_measures(model: SemanticModel, table_name: str) -> list[Measure]:
    """Order measures so the most business-relevant come first."""
    priority = re.compile(
        r"(revenue|sales|amount|value|profit|margin|cost|spend|qty|quantity|volume|units)", re.I
    )
    ms = [m for m in model.measures if m.table == table_name]
    ms.sort(key=lambda m: (0 if priority.search(m.name) else 1, "Count" in m.name, m.name))
    return ms


def _has_date_table(model: SemanticModel) -> bool:
    return any(t.is_date_table for t in model.tables)


def trend_axis(model: SemanticModel, table_name: str) -> tuple[str, str] | None:
    """Pick the column a trend chart should run along.

    A real calendar table is always best. Failing that, an ordered period label
    left behind by unpivoting a crosstab still gives a meaningful time axis —
    which matters, because a crosstab is exactly the shape that has no dates.
    """
    if _has_date_table(model):
        return (DATE_TABLE_NAME, "Month Year")
    t = model.table(table_name)
    if t is not None:
        for c in t.columns:
            if c.name == "Period" or (c.sort_by and re.search(r"period|month|week", c.name, re.I)):
                return (table_name, c.name)
    return None


# ---------------------------------------------------------------------------
# Layout — turns a page intent into positioned visuals
# ---------------------------------------------------------------------------


def _title(text: str) -> VisualSpec:
    return VisualSpec(kind="title", x=24, y=16, width=900, height=36, text=text)


def layout_overview(
    title: str,
    kpis: list[str],
    trend_category: tuple[str, str] | None,
    trend_measure: str | None,
    bar_category: tuple[str, str] | None,
    bar_measure: str | None,
    slicers: list[tuple[str, str]],
    matrix_rows: tuple[str, str] | None,
    matrix_measures: list[str],
) -> PageSpec:
    v: list[VisualSpec] = [_title(title)]

    for i, (t, c) in enumerate(slicers[:3]):
        v.append(VisualSpec(kind="slicer", x=24 + i * 248, y=64, width=232, height=80,
                            title=c, category=(t, c)))

    if kpis:
        v.append(VisualSpec(kind="card", x=24, y=156, width=1232, height=112,
                            title="Key figures", measures=kpis[:4]))

    if trend_category and trend_measure:
        v.append(VisualSpec(kind="line", x=24, y=284, width=768, height=200,
                            title=f"{trend_measure} over time",
                            measures=[trend_measure], category=trend_category))
    if bar_category and bar_measure:
        v.append(VisualSpec(kind="bar", x=808, y=284, width=448, height=200,
                            title=f"{bar_measure} by {bar_category[1]}",
                            measures=[bar_measure], category=bar_category))
    if matrix_rows and matrix_measures:
        v.append(VisualSpec(kind="matrix", x=24, y=500, width=1232, height=196,
                            title=f"{matrix_rows[1]} breakdown",
                            measures=matrix_measures[:4], category=matrix_rows))
    return PageSpec(title=title, visuals=v)


def layout_breakdown(
    title: str,
    charts: list[tuple[tuple[str, str], str]],
    matrix_rows: tuple[str, str] | None,
    matrix_series: tuple[str, str] | None,
    matrix_measures: list[str],
) -> PageSpec:
    v: list[VisualSpec] = [_title(title)]
    slots = [(24, 64, 610, 310), (654, 64, 602, 310)]
    for (cat, meas), (x, y, w, h) in zip(charts[:2], slots):
        v.append(VisualSpec(kind="column", x=x, y=y, width=w, height=h,
                            title=f"{meas} by {cat[1]}", measures=[meas], category=cat))
    if matrix_rows and matrix_measures:
        v.append(VisualSpec(kind="matrix", x=24, y=390, width=1232, height=306,
                            title="Detail matrix", measures=matrix_measures[:4],
                            category=matrix_rows, series=matrix_series))
    return PageSpec(title=title, visuals=v)


def layout_detail(title: str, columns: list[tuple[str, str]], slicers: list[tuple[str, str]],
                  measures: list[str]) -> PageSpec:
    v: list[VisualSpec] = [_title(title)]
    for i, (t, c) in enumerate(slicers[:4]):
        v.append(VisualSpec(kind="slicer", x=24 + i * 248, y=64, width=232, height=80,
                            title=c, category=(t, c)))
    v.append(VisualSpec(kind="table", x=24, y=156, width=1232, height=540,
                        title="Row-level detail", columns=columns[:12], measures=measures[:3]))
    return PageSpec(title=title, visuals=v)


# ---------------------------------------------------------------------------
# Rule-based designer
# ---------------------------------------------------------------------------


def rule_based_design(model: SemanticModel, report_title: str = "") -> ReportSpec:
    primary = model.primary_table or (model.tables[0].name if model.tables else "")
    table = model.table(primary)
    if table is None:
        return ReportSpec(title=report_title or "BizViz Dashboard")

    measures = rank_measures(model, primary)
    measure_names = [m.name for m in measures]
    dims = rank_dimensions(model, primary)

    date_cat = trend_axis(model, primary)
    top_measure = measure_names[0] if measure_names else None

    slicers = [(t, c) for t, c, _ in dims[:3]]
    if _has_date_table(model):
        slicers = [(DATE_TABLE_NAME, "Year")] + slicers[:2]

    bar_cat = (dims[0][0], dims[0][1]) if dims else None
    matrix_rows = (dims[1][0], dims[1][1]) if len(dims) > 1 else bar_cat

    title = report_title or f"{primary} Dashboard"
    pages = [
        layout_overview(
            title="Overview",
            kpis=measure_names[:4],
            trend_category=date_cat,
            trend_measure=top_measure,
            bar_category=bar_cat,
            bar_measure=top_measure,
            slicers=slicers,
            matrix_rows=matrix_rows,
            matrix_measures=measure_names[:4],
        )
    ]

    if len(dims) >= 2 and top_measure:
        charts = [((dims[0][0], dims[0][1]), top_measure)]
        if len(dims) > 1:
            second = measure_names[1] if len(measure_names) > 1 else top_measure
            charts.append(((dims[1][0], dims[1][1]), second))
        pages.append(
            layout_breakdown(
                title="Breakdown",
                charts=charts,
                matrix_rows=(dims[0][0], dims[0][1]),
                matrix_series=(dims[1][0], dims[1][1]) if len(dims) > 1 else None,
                matrix_measures=measure_names[:3],
            )
        )

    detail_cols = [
        (primary, c.name)
        for c in table.columns
        if c.role in ("dimension", "date", "key")
    ][:8]
    if detail_cols:
        pages.append(
            layout_detail("Detail", detail_cols, slicers, measure_names[:3])
        )

    return ReportSpec(title=title, pages=pages, designed_by="rules")


# ---------------------------------------------------------------------------
# Claude designer
# ---------------------------------------------------------------------------

DESIGNER_SYSTEM = """You are a Power BI dashboard designer working for a business \
consultant. You are given the STRUCTURE of a semantic model built from a client's \
Excel file — table names, column names, data types, distinct-value counts and \
value ranges. You never see the underlying rows, and you must not ask for them.

Your job is to decide what the dashboard should show. Return ONLY a JSON object, \
no prose, no code fences, matching this shape:

{
  "report_title": "short business title for the whole report",
  "narrative": "2-4 sentences: what this data appears to be about and what the \
dashboard is designed to answer. Write for a client, not an engineer.",
  "kpi_measures": ["measure name", ...],        // 2-4 existing measure names, most important first
  "trend_measure": "measure name",              // the one measure worth trending over time
  "pages": [
    {
      "title": "Overview",
      "purpose": "one line",
      "bar_category": {"table": "...", "column": "..."},
      "bar_measure": "measure name",
      "matrix_rows": {"table": "...", "column": "..."},
      "matrix_measures": ["measure name", ...],
      "slicers": [{"table": "...", "column": "..."}]
    }
  ],
  "renames": {"existing measure name": "Better Business Name"}
}

Hard rules:
- Every measure name you use MUST be one of the measure names given to you, \
spelled identically. Do not invent measures.
- Every table/column pair MUST exist in the model as given.
- Choose categories with a sensible number of distinct values (roughly 3-25). \
Never chart by a column with hundreds of distinct values.
- Never use an ID or key column as a chart category or slicer.
- Prefer 2 or 3 pages: an Overview, a Breakdown, and a Detail page.
- In "renames", give measures the name a client would use on a slide \
(e.g. "Total Order Amount" -> "Order Value"). Only rename where it genuinely \
reads better; an empty object is a valid answer."""


def _coerce_pair(d, model: SemanticModel) -> tuple[str, str] | None:
    """Validate a {'table','column'} the model actually returned."""
    if not isinstance(d, dict):
        return None
    t, c = d.get("table"), d.get("column")
    table = model.table(t) if t else None
    if table is None:
        return None
    col = next((x for x in table.columns if x.name == c), None)
    if col is None or col.role in ("key", "ignored"):
        return None
    return (t, c)


def claude_design(
    model: SemanticModel,
    api_key: str,
    model_name: str = "claude-sonnet-5",
    report_title: str = "",
) -> ReportSpec:
    """Ask Claude for a dashboard plan, then validate every reference it returns.

    Anything Claude names that does not exist in the model is dropped and the
    rule-based choice is used instead. The model gets to influence the design,
    never to introduce a field that isn't there.
    """
    import anthropic

    client = anthropic.Anthropic(api_key=api_key)
    payload = json.dumps(model.to_prompt_dict(), indent=1, default=str)
    # Keep the prompt bounded on very wide models.
    if len(payload) > 60_000:
        payload = payload[:60_000] + "\n... (truncated)"

    response = client.messages.create(
        model=model_name,
        max_tokens=2000,
        system=DESIGNER_SYSTEM,
        messages=[{"role": "user", "content": f"Here is the semantic model:\n{payload}"}],
    )
    text = response.content[0].text.strip()
    text = re.sub(r"^```(?:json)?|```$", "", text, flags=re.MULTILINE).strip()
    plan = json.loads(text)

    # Start from the rule-based design and let a *validated* plan override it.
    fallback = rule_based_design(model, report_title)
    primary = model.primary_table
    valid_measures = {m.name for m in model.measures}

    def keep_measures(names) -> list[str]:
        if not isinstance(names, list):
            return []
        return [n for n in names if isinstance(n, str) and n in valid_measures]

    kpis = keep_measures(plan.get("kpi_measures"))
    if not kpis:
        # Fall back to whatever the rule-based designer chose for the KPI row.
        kpis = next(
            (v.measures for p in fallback.pages for v in p.visuals if v.kind == "card"),
            [m.name for m in rank_measures(model, primary)][:4],
        )
    trend_measure = plan.get("trend_measure")
    if trend_measure not in valid_measures:
        trend_measure = kpis[0] if kpis else None

    dims = rank_dimensions(model, primary)
    date_cat = trend_axis(model, primary)

    pages: list[PageSpec] = []
    plan_pages = plan.get("pages") if isinstance(plan.get("pages"), list) else []

    for i, p in enumerate(plan_pages[:4]):
        if not isinstance(p, dict):
            continue
        ptitle = str(p.get("title") or f"Page {i + 1}")[:60]
        bar_cat = _coerce_pair(p.get("bar_category"), model) or (
            (dims[0][0], dims[0][1]) if dims else None
        )
        bar_meas = p.get("bar_measure")
        if bar_meas not in valid_measures:
            bar_meas = trend_measure
        mrows = _coerce_pair(p.get("matrix_rows"), model) or bar_cat
        mmeas = keep_measures(p.get("matrix_measures")) or kpis[:3]
        slicers = [s for s in (_coerce_pair(x, model) for x in (p.get("slicers") or [])) if s]
        if not slicers:
            slicers = [(t, c) for t, c, _ in dims[:2]]
        if date_cat and len(slicers) < 3:
            slicers = [(DATE_TABLE_NAME, "Year")] + slicers

        if i == 0:
            pages.append(layout_overview(ptitle, kpis, date_cat, trend_measure,
                                         bar_cat, bar_meas, slicers, mrows, mmeas))
        elif i == 1:
            charts = [(bar_cat, bar_meas)] if bar_cat and bar_meas else []
            if len(dims) > 1 and len(kpis) > 1:
                charts.append(((dims[1][0], dims[1][1]), kpis[1]))
            pages.append(layout_breakdown(ptitle, charts, mrows,
                                          (dims[1][0], dims[1][1]) if len(dims) > 1 else None,
                                          mmeas))
        else:
            table = model.table(primary)
            cols = [(primary, c.name) for c in (table.columns if table else [])
                    if c.role in ("dimension", "date")][:8]
            pages.append(layout_detail(ptitle, cols, slicers, kpis[:3]))

    if not pages:
        return fallback

    spec = ReportSpec(
        title=str(plan.get("report_title") or fallback.title)[:80],
        pages=pages,
        narrative=str(plan.get("narrative") or "")[:1200],
        designed_by="claude",
    )

    # Apply renames to the model's measures so the whole report picks them up.
    renames = plan.get("renames")
    if isinstance(renames, dict):
        apply_renames(model, spec, renames)

    return spec


def apply_renames(model: SemanticModel, spec: ReportSpec, renames: dict) -> None:
    """Rename measures in the model and everywhere the spec references them."""
    existing = {m.name for m in model.measures}
    safe: dict[str, str] = {}
    for old, new in renames.items():
        if not isinstance(old, str) or not isinstance(new, str):
            continue
        new = re.sub(r"[\[\]\"']", "", new).strip()[:60]
        if old in existing and new and new not in existing and new not in safe.values():
            safe[old] = new

    if not safe:
        return

    for m in model.measures:
        if m.name in safe:
            m.name = safe[m.name]
    for page in spec.pages:
        for v in page.visuals:
            v.measures = [safe.get(n, n) for n in v.measures]
            if v.title:
                for old, new in safe.items():
                    v.title = v.title.replace(old, new)
