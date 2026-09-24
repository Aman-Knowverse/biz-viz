"""
Stage 4 — assemble the Power BI project folder and zip it.

Two shapes come out of here.

**Living dashboard** (the default) ships a workbook the user keeps filling in,
with the model reading from it through Power Query. Add rows, hit Refresh, the
dashboard moves:

    <Project>/
      <Project>.pbip
      <Project>.SemanticModel/
      <Project>.Report/
      Data/<Project>_Data.xlsx
      START_HERE.txt

**Snapshot** embeds the data inside the model instead. Nothing to point at and
nothing to refresh — the right thing to hand a client who just needs to open it.
"""

from __future__ import annotations

import io
import os
import re
import tempfile
import zipfile
from dataclasses import dataclass, field
from pathlib import Path

from .design import ReportSpec
from .excel_writer import SheetPlan, write_master_workbook
from .modeling import SemanticModel
from .pbir import render_pbip_manifest, render_report
from .tmdl import render_semantic_model, should_inline

LINKED = "linked"
SNAPSHOT = "snapshot"


@dataclass
class BuildResult:
    project_name: str
    files: dict[str, str | bytes]  # path relative to the zip root
    mode: str = SNAPSHOT
    inline_data: bool = True
    data_folder_hint: str = ""
    workbook_name: str = ""
    sheet_plans: list[SheetPlan] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)

    def to_zip_bytes(self) -> bytes:
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
            for path, content in self.files.items():
                z.writestr(path, content if isinstance(content, bytes) else content.encode("utf-8"))
        return buf.getvalue()

    def write_to(self, directory: str | Path) -> Path:
        root = Path(directory)
        for path, content in self.files.items():
            target = root / path
            target.parent.mkdir(parents=True, exist_ok=True)
            if isinstance(content, bytes):
                target.write_bytes(content)
            else:
                target.write_text(content, encoding="utf-8")
        return root / self.project_name


def safe_project_name(name: str) -> str:
    s = re.sub(r'[<>:"/\\|?*]+', " ", str(name))
    s = re.sub(r"\s+", " ", s).strip()
    return (s or "BizViz Dashboard")[:60]


def _file_stem(name: str) -> str:
    return re.sub(r"\s+", "_", safe_project_name(name))


# ---------------------------------------------------------------------------
# The instruction file
# ---------------------------------------------------------------------------

START_HERE_LINKED = """\
{project}
{underline}

WHAT YOU HAVE HERE
------------------
  {project}.pbip           <- the dashboard. Double-click to open in Power BI Desktop.
  {project}.SemanticModel\\  <- how the data is joined up and calculated
  {project}.Report\\         <- the pages and charts ({n_pages} pages, {n_visuals} visuals)
  Data\\{workbook}
                            <- YOUR DATA. This is the file you keep filling in.

The dashboard reads from that Excel file. Add rows to it, hit Refresh in Power
BI, and every chart updates. You never need to come back to Biz-Viz.


STEP 1 - EXTRACT IT TO THE RIGHT PLACE
--------------------------------------
Power BI has to be told exactly where the Excel file lives - it cannot work it
out for itself. This project was built expecting:

    {data_folder}\\{workbook}

So: right-click the zip, Extract All, and extract it into

    {extract_root}

That puts the workbook exactly where the dashboard expects it. (Do not just
double-click the zip and open the files from inside it - Windows only pretends
that is a folder, and the dashboard will not find its data.)

Extracted somewhere else? Not a problem, it is a 20-second fix - see
"IF YOU MOVED THE FILES" at the bottom.


STEP 2 - OPEN IT
----------------
  1. Open Power BI Desktop.
  2. File > Open > Browse.
  3. Pick {project}.pbip
  4. It loads the data and opens the report.

If Power BI says the project format is a preview feature:
  File > Options and settings > Options > Preview features,
  tick anything mentioning "Power BI Project (.pbip)" or
  "enhanced report format (PBIR)", restart Power BI, try again.


STEP 3 - WANT AN ORDINARY .PBIX FILE?
-------------------------------------
Once it is open:  File > Save As > Power BI files (*.pbix)

Keep the Excel file where it is. The .pbix still reads from it, so refresh
keeps working exactly the same way.


HOW TO ADD NEW DATA FROM NOW ON
-------------------------------
This is the part that matters. From here on, {workbook} is your
master data file. Do not go back to your original spreadsheet.

  1. Open  Data\\{workbook}
  2. Go to the sheet you want to add to.
  3. Type into the first empty row inside the coloured table. Blank rows are
     already sitting there waiting for you - you do not need to insert rows.
  4. Save and close the file.
  5. In Power BI, click Home > Refresh.

That is it. New rows, new months, new categories - all of it flows through.

FOUR THINGS THAT WILL BREAK IT
  - Renaming a column heading. The dashboard finds columns by name.
    (The heading row is locked to make this hard to do by accident.)
  - Inserting or deleting a column in the middle of a table.
  - Typing a "Total" row. Power BI works out its own totals; a typed one
    gets counted twice.
  - Typing outside the coloured table. Anything beside or below it is ignored.

There is a "READ ME FIRST" sheet inside the workbook saying the same thing,
for whoever ends up maintaining it after you.


WHAT WAS DONE TO YOUR ORIGINAL FILE
-----------------------------------
{cleaning}
{reshaping}

WHAT THE DASHBOARD CALCULATES
-----------------------------
{measures}

IF YOU MOVED THE FILES
----------------------
The folder is stored as a setting you can edit, not buried in code:

  In Power BI:  Home > Transform data > Manage Parameters
  Set "DataFolder" to the folder that contains {workbook}
  Then:  Home > Close & Apply

Still stuck? File > Options and settings > Data source settings > Change Source
does the same job.


ONE HONEST WARNING
------------------
The way your sheets were joined together, and the calculations written for you,
were worked out automatically from column names and values. That is a good first
draft, not gospel. Before this goes in front of anyone who matters, open the
model view in Power BI and check the lines between the tables look right, and
sanity-check one or two headline numbers against your own reporting.
"""

START_HERE_SNAPSHOT = """\
{project}
{underline}

WHAT YOU HAVE HERE
------------------
  {project}.pbip           <- the dashboard. Double-click to open in Power BI Desktop.
  {project}.SemanticModel\\  <- the data model, with your data held inside it
  {project}.Report\\         <- the pages and charts ({n_pages} pages, {n_visuals} visuals)

This is a SNAPSHOT. Your data is stored inside the file itself, so it opens
anywhere with nothing to set up and no paths to fix - which is what makes it
the right thing to email to somebody.

The trade-off: it cannot be refreshed. It shows the data as it was on the day
it was generated. If you need a dashboard that keeps up with new data, generate
it again from Biz-Viz using the "Living dashboard" option instead.


TO OPEN IT
----------
  1. Extract the whole zip to a folder (keep the folders together).
  2. Open Power BI Desktop.
  3. File > Open > Browse, and pick {project}.pbip

If Power BI says the project format is a preview feature:
  File > Options and settings > Options > Preview features,
  tick anything mentioning "Power BI Project (.pbip)" or
  "enhanced report format (PBIR)", restart Power BI, try again.

For an ordinary .pbix:  File > Save As > Power BI files (*.pbix)


WHAT WAS DONE TO YOUR ORIGINAL FILE
-----------------------------------
{cleaning}

WHAT THE DASHBOARD CALCULATES
-----------------------------
{measures}


ONE HONEST WARNING
------------------
The way your sheets were joined together, and the calculations written for you,
were worked out automatically from column names and values. That is a good first
draft, not gospel. Check the model view and a couple of headline numbers before
this goes in front of anyone who matters.
"""


def _cleaning_notes(model: SemanticModel) -> str:
    lines = [f"  [{t.name}] {note}" for t in model.tables for note in t.notes]
    return "\n".join(lines) if lines else "  Your sheets were already tidy - nothing needed restructuring."


def _reshaping_notes(model: SemanticModel) -> str:
    out = []
    for t in model.tables:
        if t.unpivot_plan is None:
            continue
        p = t.unpivot_plan
        shown = ", ".join(str(c) for c in p.period_columns[:5])
        more = "..." if len(p.period_columns) > 5 else ""
        out.append(
            f"\n  '{t.name}' was left in its original side-by-side layout on purpose.\n"
            f"  Its {len(p.period_columns)} period columns ({shown}{more}) are turned into rows\n"
            f"  by Power BI every time it refreshes - so you carry on filling it across,\n"
            f"  the way you always have, and Power BI does the rearranging."
        )
    return "\n".join(out)


def _measure_notes(model: SemanticModel) -> str:
    lines = [f"  {m.name} = {m.dax}" for m in model.measures[:40]]
    if len(model.measures) > 40:
        lines.append(f"  ...and {len(model.measures) - 40} more")
    return "\n".join(lines) if lines else "  none"


# ---------------------------------------------------------------------------
# Build
# ---------------------------------------------------------------------------


def build_project(
    model: SemanticModel,
    spec: ReportSpec,
    project_name: str = "",
    mode: str = LINKED,
    extract_root: str = r"C:\BizViz",
    guard_rails: bool = True,
    data_folder: str | None = None,  # kept for older callers
) -> BuildResult:
    name = safe_project_name(project_name or spec.title or "BizViz Dashboard")
    sm_folder = f"{name}.SemanticModel"
    rp_folder = f"{name}.Report"

    files: dict[str, str | bytes] = {}
    warnings: list[str] = []
    plans: list[SheetPlan] = []
    workbook_name = ""
    excel_tables: dict[str, str] = {}

    extract_root = (extract_root or r"C:\BizViz").rstrip("\\/")
    resolved_data_folder = data_folder or f"{extract_root}\\{name}\\Data"

    if mode == LINKED:
        workbook_name = f"{_file_stem(name)}_Data.xlsx"
        # openpyxl only writes to a real path, so build it in a temp file and
        # carry the bytes into the zip.
        tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
        tmp.close()
        try:
            plans = write_master_workbook(
                model.tables, tmp.name, project_name=name, guard_rails=guard_rails
            )
            files[f"{name}/Data/{workbook_name}"] = Path(tmp.name).read_bytes()
        finally:
            os.unlink(tmp.name)

        excel_tables = {p.table_name: p.excel_table_name for p in plans}
        if not excel_tables:
            warnings.append(
                "No sheet could be written to the workbook, so the project falls back to "
                "holding its data internally. It will open, but it will not refresh."
            )
            mode = SNAPSHOT
            files.pop(f"{name}/Data/{workbook_name}", None)
            workbook_name = ""
        else:
            unprotected = [p.sheet_name for p in plans if not p.protected]
            if unprotected and guard_rails:
                warnings.append(
                    "Headings were left unlocked on "
                    + ", ".join(f"'{s}'" for s in unprotected)
                    + " because the sheet is too large for cell-level protection to be worth "
                    "it. A sheet that size is normally a system export rather than something "
                    "typed by hand."
                )

    inline = should_inline(model)
    if mode == SNAPSHOT and not inline:
        warnings.append(
            "The data is too large to embed, so it ships as CSV in a Data folder and the "
            f"model expects it at {resolved_data_folder}. The 'Living dashboard' option "
            "handles large data more gracefully."
        )

    for rel, content in render_semantic_model(
        model,
        default_data_folder=resolved_data_folder,
        mode=mode,
        workbook_file=workbook_name,
        excel_tables=excel_tables,
    ).items():
        files[f"{name}/{sm_folder}/{rel}"] = content

    for rel, content in render_report(spec, model, sm_folder).items():
        files[f"{name}/{rp_folder}/{rel}"] = content

    files[f"{name}/{name}.pbip"] = render_pbip_manifest(rp_folder)

    if mode == SNAPSHOT and not inline:
        for t in model.tables:
            if t.is_date_table:
                continue
            buf = io.StringIO()
            t.frame.to_csv(buf, index=False)
            files[f"{name}/Data/{t.name}.csv"] = buf.getvalue()

    n_visuals = sum(len(p.visuals) for p in spec.pages)
    template = START_HERE_LINKED if mode == LINKED else START_HERE_SNAPSHOT
    files[f"{name}/START_HERE.txt"] = template.format(
        project=name,
        underline="=" * len(name),
        n_pages=len(spec.pages),
        n_visuals=n_visuals,
        workbook=workbook_name,
        data_folder=resolved_data_folder,
        extract_root=extract_root,
        cleaning=_cleaning_notes(model),
        reshaping=_reshaping_notes(model),
        measures=_measure_notes(model),
    )

    if spec.narrative:
        files[f"{name}/DASHBOARD_NOTES.txt"] = (
            "What this dashboard is for\n"
            "==========================\n\n" + spec.narrative + "\n"
        )

    return BuildResult(
        project_name=name,
        files=files,
        mode=mode,
        inline_data=inline if mode == SNAPSHOT else False,
        data_folder_hint=resolved_data_folder,
        workbook_name=workbook_name,
        sheet_plans=plans,
        warnings=warnings,
    )
