"""
Stage 3c — write the report as PBIR.

PBIR is the text form of a Power BI report: one JSON file per visual, arranged
in folders under `definition/`. Everything here follows Microsoft's published
PBIR contracts; the shapes that are easy to get subtly wrong (role names,
where `queryState` sits, how a projection references a measure) are centralised
in the builder functions rather than repeated per visual type.
"""

from __future__ import annotations

import json
import secrets

from .design import PageSpec, ReportSpec, VisualSpec
from .modeling import SemanticModel
from .tmdl import platform_file

SCHEMA = "https://developer.microsoft.com/json-schemas/fabric/item/report/definition"
VISUAL_SCHEMA = f"{SCHEMA}/visualContainer/2.9.0/schema.json"
PAGE_SCHEMA = f"{SCHEMA}/page/2.1.0/schema.json"
PAGES_SCHEMA = f"{SCHEMA}/pagesMetadata/1.0.0/schema.json"
REPORT_SCHEMA = f"{SCHEMA}/report/3.3.0/schema.json"
VERSION_SCHEMA = f"{SCHEMA}/versionMetadata/1.0.0/schema.json"
PBIR_SCHEMA = "https://developer.microsoft.com/json-schemas/fabric/item/report/definitionProperties/2.0.0/schema.json"
PBIP_SCHEMA = "https://developer.microsoft.com/json-schemas/fabric/pbip/pbipProperties/1.0.0/schema.json"


# ---------------------------------------------------------------------------
# Identifiers
# ---------------------------------------------------------------------------


def visual_id() -> str:
    return secrets.token_hex(10)  # 20 hex chars


def page_id() -> str:
    return "ReportSection" + secrets.token_hex(12)  # ReportSection + 24 hex


# ---------------------------------------------------------------------------
# Expression trees
# ---------------------------------------------------------------------------


def column_field(table: str, column: str) -> dict:
    return {"Column": {"Expression": {"SourceRef": {"Entity": table}}, "Property": column}}


def measure_field(table: str, measure: str) -> dict:
    return {"Measure": {"Expression": {"SourceRef": {"Entity": table}}, "Property": measure}}


def column_projection(table: str, column: str, active: bool = False) -> dict:
    p = {
        "field": column_field(table, column),
        "queryRef": f"{table}.{column}",
        "nativeQueryRef": column,
    }
    if active:
        p["active"] = True
    return p


def measure_projection(table: str, measure: str) -> dict:
    return {
        "field": measure_field(table, measure),
        "queryRef": f"{table}.{measure}",
        "nativeQueryRef": measure,
    }


AGG_FUNCTION = {"SUM": 0, "AVERAGE": 1, "COUNT": 2, "MIN": 3, "MAX": 4, "COUNTROWS": 5}
AGG_QUERYREF = {"SUM": "Sum", "AVERAGE": "Average", "COUNT": "Count",
                "MIN": "Min", "MAX": "Max", "COUNTROWS": "CountNonNull"}
AGG_LABEL = {"SUM": "Sum of", "AVERAGE": "Average of", "COUNT": "Count of",
             "MIN": "Min of", "MAX": "Max of", "COUNTROWS": "Count of"}


def aggregation_field(table: str, column: str, agg: str) -> dict:
    return {
        "Aggregation": {
            "Expression": {
                "Column": {
                    "Expression": {"SourceRef": {"Entity": table}},
                    "Property": column,
                }
            },
            "Function": AGG_FUNCTION.get(agg, 0),
        }
    }


def aggregation_projection(table: str, column: str, agg: str) -> dict:
    return {
        "field": aggregation_field(table, column, agg),
        "queryRef": f"{AGG_QUERYREF.get(agg, 'Sum')}({table}.{column})",
        # Aggregation projections must spell this "Sum of <Column>"; the bare
        # column name renders an empty visual with no error at all.
        "nativeQueryRef": f"{AGG_LABEL.get(agg, 'Sum of')} {column}",
    }


class ValueBinder:
    """Turns a measure name from the design into something a visual can bind to.

    Power BI proved willing to drop measures defined in the model file while
    happily keeping the columns underneath them, which left every measure-bound
    visual showing "Missing_References". So where a measure is just an
    aggregation of one column — which all the generated ones are — the visual
    binds to that column directly instead. It is exactly the expression Power BI
    writes when a person drags a numeric field onto a chart, so it cannot be
    rejected for a reason the columns themselves survive.

    The measures are still written into the model; they are simply no longer
    load-bearing.
    """

    def __init__(self, model, prefer_columns: bool = True):
        self.prefer_columns = prefer_columns
        self.by_name = {m.name: m for m in model.measures}

    def __contains__(self, name: str) -> bool:
        return name in self.by_name

    def field(self, name: str) -> dict | None:
        m = self.by_name.get(name)
        if m is None:
            return None
        if self.prefer_columns and m.source_column and m.agg:
            return aggregation_field(m.table, m.source_column, m.agg)
        return measure_field(m.table, name)

    def projection(self, name: str) -> dict | None:
        m = self.by_name.get(name)
        if m is None:
            return None
        if self.prefer_columns and m.source_column and m.agg:
            return aggregation_projection(m.table, m.source_column, m.agg)
        return measure_projection(m.table, name)

    def projections(self, names) -> list[dict]:
        out = [self.projection(n) for n in names]
        return [p for p in out if p]


def lit(value) -> dict:
    """A PBIR literal expression. Booleans and numbers are unquoted; text is."""
    if isinstance(value, bool):
        v = "true" if value else "false"
    elif isinstance(value, (int, float)):
        v = str(value)
    else:
        v = f"'{value}'"
    return {"expr": {"Literal": {"Value": v}}}


def position(v: VisualSpec, z: int) -> dict:
    return {
        "x": v.x, "y": v.y, "z": z,
        "height": v.height, "width": v.width,
        "tabOrder": z,
    }


# ---------------------------------------------------------------------------
# Visual builders
# ---------------------------------------------------------------------------


def title_vco(title: str) -> dict:
    """A visual's header title.

    `title` is a *container* object, not a visual object — it belongs under
    `visualContainerObjects`, alongside background and border. Putting it in
    `objects` validates but never renders.
    """
    if not title:
        return {}
    return {
        "visualContainerObjects": {
            "title": [{
                "properties": {
                    "show": lit(True),
                    "text": lit(title),
                    "fontSize": lit(11),
                }
            }]
        }
    }


def build_title(spec: VisualSpec, z: int) -> dict:
    return {
        "$schema": VISUAL_SCHEMA,
        "name": visual_id(),
        "position": position(spec, z),
        "visual": {
            "visualType": "textbox",
            "objects": {
                "general": [{
                    "properties": {
                        "paragraphs": [{
                            "textRuns": [{
                                "value": spec.text,
                                "textStyle": {"fontFamily": "Segoe UI Semibold", "fontSize": "20px"},
                            }],
                            "horizontalTextAlignment": "left",
                        }]
                    }
                }]
            },
            "visualContainerObjects": {
                "background": [{"properties": {"show": lit(False)}}],
                "border": [{"properties": {"show": lit(False)}}],
                "padding": [{
                    "properties": {
                        "top": {"expr": {"Literal": {"Value": "0D"}}},
                        "bottom": {"expr": {"Literal": {"Value": "0D"}}},
                        "left": {"expr": {"Literal": {"Value": "0D"}}},
                        "right": {"expr": {"Literal": {"Value": "0D"}}},
                    }
                }],
            },
        },
    }


def build_card(spec: VisualSpec, z: int, binder: "ValueBinder") -> dict:
    projections = binder.projections(spec.measures)
    if not projections:
        return {}
    return {
        "$schema": VISUAL_SCHEMA,
        "name": visual_id(),
        "position": position(spec, z),
        "visual": {
            "visualType": "cardVisual",
            "query": {"queryState": {"Data": {"projections": projections}}},
            "objects": {
                # The internal rectangle around each callout is noise on a KPI row.
                "outline": [{"properties": {"show": lit(False)}, "selector": {"id": "default"}}],
            },
        },
    }


def build_cartesian(spec: VisualSpec, z: int, binder: "ValueBinder",
                    visual_type: str) -> dict:
    if not spec.category or not spec.measures:
        return {}
    cat_table, cat_col = spec.category
    y_projections = binder.projections(spec.measures)
    if not y_projections:
        return {}

    query: dict = {
        "queryState": {
            "Category": {"projections": [column_projection(cat_table, cat_col, active=True)]},
            "Y": {"projections": y_projections},
        }
    }
    # Bar/column charts read best sorted by value; a time axis must stay in
    # chronological order, so only sort the non-time charts.
    if visual_type in ("barChart", "clusteredBarChart", "columnChart", "clusteredColumnChart"):
        sort_field = binder.field(spec.measures[0])
        if sort_field:
            query["sortDefinition"] = {
                "sort": [{"field": sort_field, "direction": "Descending"}],
                "isDefaultSort": False,
            }

    return {
        "$schema": VISUAL_SCHEMA,
        "name": visual_id(),
        "position": position(spec, z),
        "visual": {
            "visualType": visual_type,
            "query": query,
            "objects": {"labels": [{"properties": {"show": lit(False)}}]},
            **title_vco(spec.title),
        },
    }


def build_matrix(spec: VisualSpec, z: int, binder: "ValueBinder") -> dict:
    if not spec.category:
        return {}
    rows_table, rows_col = spec.category
    values = binder.projections(spec.measures)
    if not values:
        return {}

    query_state: dict = {
        "Rows": {"projections": [column_projection(rows_table, rows_col)]},
        "Values": {"projections": values},
    }
    if spec.series:
        s_table, s_col = spec.series
        query_state["Columns"] = {"projections": [column_projection(s_table, s_col)]}

    return {
        "$schema": VISUAL_SCHEMA,
        "name": visual_id(),
        "position": position(spec, z),
        "visual": {
            "visualType": "pivotTable",
            "query": {"queryState": query_state},
            "objects": {
                "columnHeaders": [{
                    "properties": {
                        "columnAdjustment": lit("growToFit"),
                        "autoSizeColumnWidth": lit(True),
                    }
                }]
            },
            **title_vco(spec.title),
        },
    }


def build_table(spec: VisualSpec, z: int, binder: "ValueBinder") -> dict:
    projections = [column_projection(t, c) for t, c in spec.columns]
    projections += binder.projections(spec.measures)
    if not projections:
        return {}
    return {
        "$schema": VISUAL_SCHEMA,
        "name": visual_id(),
        "position": position(spec, z),
        "visual": {
            "visualType": "tableEx",
            "query": {"queryState": {"Values": {"projections": projections}}},
            "objects": {
                "columnHeaders": [{
                    "properties": {
                        "columnAdjustment": lit("growToFit"),
                        "autoSizeColumnWidth": lit(True),
                    }
                }]
            },
            **title_vco(spec.title),
        },
    }


def build_slicer(spec: VisualSpec, z: int) -> dict:
    if not spec.category:
        return {}
    t, c = spec.category
    return {
        "$schema": VISUAL_SCHEMA,
        "name": visual_id(),
        "position": position(spec, z),
        "visual": {
            "visualType": "slicer",
            "query": {"queryState": {"Values": {"projections": [column_projection(t, c)]}}},
            "objects": {
                "data": [{"properties": {"mode": lit("Dropdown")}}],
                "header": [{"properties": {"show": lit(True), "text": lit(c)}}],
            },
        },
    }


BUILDERS = {
    "title": lambda s, z, b: build_title(s, z),
    "card": build_card,
    "line": lambda s, z, b: build_cartesian(s, z, b, "lineChart"),
    "bar": lambda s, z, b: build_cartesian(s, z, b, "barChart"),
    "column": lambda s, z, b: build_cartesian(s, z, b, "clusteredColumnChart"),
    "matrix": build_matrix,
    "table": build_table,
    "slicer": lambda s, z, b: build_slicer(s, z),
}


# ---------------------------------------------------------------------------
# Report assembly
# ---------------------------------------------------------------------------


def render_report(
    spec: ReportSpec,
    model: SemanticModel,
    semantic_model_folder_name: str,
    bind_to_columns: bool = True,
) -> dict[str, str]:
    """Return {relative path: file content} for the .Report folder."""
    binder = ValueBinder(model, prefer_columns=bind_to_columns)

    files: dict[str, str] = {}
    page_names: list[str] = []
    first_page: str | None = None

    for page in spec.pages:
        pname = page_id()
        page_names.append(pname)
        if first_page is None:
            first_page = pname

        z = 1000
        for v in page.visuals:
            builder = BUILDERS.get(v.kind)
            if builder is None:
                continue
            payload = builder(v, z, binder)
            if not payload:
                continue  # a visual whose fields did not resolve is dropped, not faked
            vname = payload["name"]
            files[f"definition/pages/{pname}/visuals/{vname}/visual.json"] = _dump(payload)
            z += 1000

        files[f"definition/pages/{pname}/page.json"] = _dump({
            "$schema": PAGE_SCHEMA,
            "name": pname,
            "displayName": page.title,
            "displayOption": "FitToPage",
            "height": 720,
            "width": 1280,
        })

    files["definition/pages/pages.json"] = _dump({
        "$schema": PAGES_SCHEMA,
        "pageOrder": page_names,
        "activePageName": first_page or "",
    })

    # `reportVersionAtImport` is required and must match the schema versions we
    # actually wrote above, or Desktop applies the theme against the wrong
    # contract. `layoutOptimization` is deliberately absent: it is not a
    # property of report.json in the 3.x contract.
    files["definition/report.json"] = _dump({
        "$schema": REPORT_SCHEMA,
        "themeCollection": {
            "baseTheme": {
                "name": "CY24SU10",
                "reportVersionAtImport": {
                    "visual": "2.9.0",
                    "page": "2.1.0",
                    "report": "3.3.0",
                },
                "type": "SharedResources",
            }
        },
    })

    files["definition/version.json"] = _dump({"$schema": VERSION_SCHEMA, "version": "2.0.0"})

    files["definition.pbir"] = _dump({
        "$schema": PBIR_SCHEMA,
        "version": "4.0",
        "datasetReference": {"byPath": {"path": f"../{semantic_model_folder_name}"}},
    })

    files[".platform"] = platform_file("Report", spec.title)
    return files


def render_pbip_manifest(report_folder_name: str) -> str:
    return _dump({
        "$schema": PBIP_SCHEMA,
        "version": "1.0",
        "artifacts": [{"report": {"path": report_folder_name}}],
        "settings": {"enableAutoRecovery": True},
    })


def _dump(obj: dict) -> str:
    return json.dumps(obj, indent=2, ensure_ascii=False)
