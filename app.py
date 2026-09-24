"""
Biz-Viz — turn a client's Excel workbook into an openable Power BI project.

Same shape as Flow Finder: upload, confirm what was detected, generate, download.
The difference is what comes out — not a diagram, but a real .pbip project that
Power BI Desktop opens with the data model and dashboard already built.
"""

import pandas as pd
import streamlit as st

from bizviz.build import LINKED, SNAPSHOT, build_project
from bizviz.design import claude_design, rule_based_design
from bizviz.modeling import build_model
from bizviz.profiling import profile_workbook

st.set_page_config(page_title="Biz-Viz — Excel to Power BI", page_icon="📊", layout="wide")

ROLES = ["measure", "dimension", "date", "key", "ignored"]
ROLE_HELP = {
    "measure": "A number to add up — revenue, quantity, hours.",
    "dimension": "A category to slice by — region, product, department.",
    "date": "A date, used for trends and the calendar table.",
    "key": "An ID used to join tables. Hidden in the report.",
    "ignored": "Left out of the model entirely.",
}

st.title("📊 Biz-Viz")
st.caption(
    "Upload a client's Excel workbook and get back a Power BI project — cleaned data, "
    "a semantic model with relationships and measures, and a starter dashboard — "
    "that opens in Power BI Desktop."
)

with st.expander("What this actually produces, and what it doesn't", expanded=False):
    st.markdown(
        """
**You get a `.pbip` project**, not a `.pbix`. A `.pbix` is an undocumented binary
container that cannot be reliably generated outside Power BI itself. A `.pbip` is
Microsoft's *project* format — the same report, stored as text — and Power BI
Desktop opens it natively. Once it is open, **File → Save As → .pbix** if you need one.

**What gets built for you**
- Messy sheets cleaned: title blocks skipped, blank and subtotal rows removed,
  merged label cells filled down, crosstabs unpivoted into rows
- A semantic model: typed columns, detected relationships, a proper calendar table
- DAX measures for every numeric column, plus row counts
- A dashboard with KPI cards, a trend, breakdowns by category, a matrix and a detail table

**What it doesn't do**
- It cannot know your business. Check the relationships and measures before you
  put this in front of a client — automatic detection is a starting point.
- It reads your workbook and writes a project. It never connects to a source
  system and never writes back anywhere.
"""
    )

# ---------------------------------------------------------------------------
# 1. Upload
# ---------------------------------------------------------------------------

st.subheader("1. Upload the workbook")
uploaded = st.file_uploader(
    "Excel or CSV", type=["xlsx", "xlsm", "xls", "csv", "tsv"],
    help="Every sheet with data is read. Sheets that are just notes or blank are skipped.",
)

if uploaded is None:
    st.info("Waiting for a file. Multi-sheet workbooks work best — related sheets become linked tables.")
    st.stop()

if st.session_state.get("_file_id") != (uploaded.name, uploaded.size):
    st.session_state.clear()
    st.session_state["_file_id"] = (uploaded.name, uploaded.size)

auto_unpivot = st.checkbox(
    "Unpivot crosstab sheets automatically",
    value=True,
    help="If a sheet has months or years as column headers, turn them into rows. "
         "Power BI needs one row per period.",
)

if "profiles" not in st.session_state or st.session_state.get("_unpivot") != auto_unpivot:
    try:
        with st.spinner("Reading and cleaning the workbook..."):
            st.session_state["profiles"] = profile_workbook(uploaded, auto_unpivot=auto_unpivot)
        st.session_state["_unpivot"] = auto_unpivot
    except Exception as e:
        st.error(f"Couldn't read that file: {e}")
        st.stop()

profiles = st.session_state["profiles"]
if not profiles:
    st.error(
        "No usable tables found. Each sheet needs a header row and at least two columns "
        "of data underneath it."
    )
    st.stop()

st.success(f"Found {len(profiles)} usable table(s).")

# ---------------------------------------------------------------------------
# 2. Review what was detected
# ---------------------------------------------------------------------------

st.subheader("2. Check what it worked out")
st.caption(
    "Every column was given a role. Change any that look wrong — this drives the whole model, "
    "so a numeric ID left as a measure will show up as a nonsense total."
)

tabs = st.tabs([f"{p.name} ({len(p.frame):,})" for p in profiles])
for tab, prof in zip(tabs, profiles):
    with tab:
        for note in prof.notes:
            st.info(note, icon="🧹")

        editor_df = pd.DataFrame(
            [
                {
                    "Column": c.name,
                    "Role": c.role,
                    "Type": c.dtype,
                    "Distinct": c.n_unique,
                    "Blank %": round(c.null_pct, 1),
                    "Why": c.reason,
                }
                for c in prof.columns
            ]
        )
        edited = st.data_editor(
            editor_df,
            key=f"editor_{prof.name}",
            hide_index=True,
            use_container_width=True,
            column_config={
                "Role": st.column_config.SelectboxColumn(
                    "Role", options=ROLES, required=True,
                    help=" · ".join(f"{k}: {v}" for k, v in ROLE_HELP.items()),
                ),
                "Why": st.column_config.TextColumn("Why it was classified this way", width="large"),
            },
            disabled=["Column", "Type", "Distinct", "Blank %", "Why"],
        )
        # Push edits back onto the profile so the model rebuild picks them up.
        for _, row in edited.iterrows():
            col = prof.column(row["Column"])
            if col is not None and col.role != row["Role"]:
                col.role = row["Role"]
                col.reason = "Set by you."

        with st.expander("Preview the cleaned data"):
            st.dataframe(prof.frame.head(8), use_container_width=True)

# ---------------------------------------------------------------------------
# 3. Build the model
# ---------------------------------------------------------------------------

st.subheader("3. Build the model")
add_date_table = st.checkbox(
    "Add a calendar table", value=True,
    help="Creates a proper Date table with Year, Quarter and Month, joined to your date columns. "
         "Needed for any time intelligence later.",
)

if st.button("Build the model", type="primary"):
    with st.spinner("Detecting relationships and writing measures..."):
        st.session_state["model"] = build_model(profiles, add_date_table=add_date_table)
    st.session_state.pop("spec", None)

model = st.session_state.get("model")
if model is None:
    st.stop()

c1, c2, c3 = st.columns(3)
c1.metric("Tables", len(model.tables))
c2.metric("Relationships", len(model.relationships))
c3.metric("Measures", len(model.measures))

left, right = st.columns(2)
with left:
    st.markdown("**Relationships found**")
    if model.relationships:
        st.dataframe(
            pd.DataFrame(
                [
                    {"From": f"{r.from_table}[{r.from_column}]",
                     "To": f"{r.to_table}[{r.to_column}]",
                     "Why": r.reason}
                    for r in model.relationships
                ]
            ),
            hide_index=True, use_container_width=True, height=240,
        )
    else:
        st.caption("None — the sheets share no matching key columns, so they stay independent tables.")

    if model.dropped_relationships:
        with st.expander(f"{len(model.dropped_relationships)} link(s) deliberately left out"):
            st.caption(
                "Power BI allows only one filter path between any two tables. These were "
                "the weaker duplicate of a link that was kept."
            )
            for d in model.dropped_relationships:
                st.markdown(f"- {d}")

with right:
    st.markdown("**Measures written**")
    st.dataframe(
        pd.DataFrame([{"Measure": m.name, "DAX": m.dax} for m in model.measures]),
        hide_index=True, use_container_width=True, height=240,
    )

# ---------------------------------------------------------------------------
# 4. Design the dashboard
# ---------------------------------------------------------------------------

st.subheader("4. Design the dashboard")
mode = st.radio(
    "How should the layout be decided?",
    ["Rules (free, offline)", "Let Claude design it (uses your API key)"],
    horizontal=True,
    help="Rules pick the biggest table, its strongest measures and categories. "
         "Claude reads the model's structure and decides what the pages should be "
         "and what to call things in client language.",
)

api_key = ""
if mode.startswith("Let Claude"):
    api_key = st.text_input(
        "Your Anthropic API key (starts with sk-ant-)", type="password",
        help="Costs a fraction of a cent per run. Only the model's structure is sent — "
             "table names, column names, types and counts. None of your data rows leave your machine.",
    )
    st.caption(
        "🔒 Claude sees column names and shapes, never the rows themselves. It returns a plan; "
        "the files are still written by code, and anything it names that doesn't exist is discarded."
    )

project_name = st.text_input("Project name", value="BizViz Dashboard")

if st.button("Design the dashboard", type="primary"):
    if mode.startswith("Let Claude") and not api_key:
        st.warning("Paste an API key, or switch to the rules option.")
    else:
        try:
            with st.spinner("Deciding what to show..."):
                if mode.startswith("Let Claude"):
                    st.session_state["spec"] = claude_design(model, api_key, report_title=project_name)
                else:
                    st.session_state["spec"] = rule_based_design(model, report_title=project_name)
        except Exception as e:
            st.error(f"Claude's design step failed ({e}). Falling back to the rules-based layout.")
            st.session_state["spec"] = rule_based_design(model, report_title=project_name)

spec = st.session_state.get("spec")
if spec is None:
    st.stop()

if spec.narrative:
    st.info(spec.narrative, icon="💬")

st.markdown(f"**{spec.title}** — {len(spec.pages)} page(s), designed by *{spec.designed_by}*")
for page in spec.pages:
    with st.expander(f"Page: {page.title}", expanded=False):
        rows = []
        for v in page.visuals:
            rows.append({
                "Visual": v.kind,
                "Shows": v.text or v.title,
                "Fields": ", ".join(
                    ([f"{v.category[1]}"] if v.category else []) + v.measures
                    + [c for _, c in v.columns]
                ) or "—",
            })
        st.dataframe(pd.DataFrame(rows), hide_index=True, use_container_width=True)

# ---------------------------------------------------------------------------
# 5. Generate
# ---------------------------------------------------------------------------

st.subheader("5. Generate the Power BI project")

mode_label = st.radio(
    "What kind of file do you need?",
    [
        "Living dashboard — comes with an Excel file you keep filling in",
        "Snapshot — everything sealed inside one file, opens anywhere",
    ],
    help=(
        "A living dashboard ships a clean Excel workbook alongside the report. Add rows to it, "
        "hit Refresh in Power BI, and the charts move. A snapshot holds the data inside the "
        "file itself — nothing to set up and no paths to fix, but it can never be refreshed."
    ),
)
mode = LINKED if mode_label.startswith("Living") else SNAPSHOT

extract_root = r"C:\BizViz"
guard_rails = True

if mode == LINKED:
    c1, c2 = st.columns([3, 2])
    with c1:
        extract_root = st.text_input(
            "Which folder will you extract the zip into?",
            value=r"C:\BizViz",
            help=(
                "Power BI has to be told exactly where the Excel file lives — it cannot work it "
                "out for itself. Extract the zip here and it just works. If you put it somewhere "
                "else, it is a one-line fix in Power BI's Manage Parameters dialog."
            ),
        )
    with c2:
        guard_rails = st.checkbox(
            "Protect the workbook", value=True,
            help=(
                "Puts dropdowns on the category columns and locks the heading row, so a typo "
                "cannot quietly add a fifth region and a renamed column cannot break the report."
            ),
        )
    st.caption(
        f"The workbook will be expected at "
        f"`{extract_root.rstrip(chr(92))}\\{project_name}\\Data\\...xlsx`. "
        "Everything is explained in the START_HERE.txt inside the zip."
    )
else:
    st.caption(
        "The data will be embedded in the file. Best for sending someone a dashboard that "
        "just opens — but it shows the data as it is today and cannot be refreshed."
    )

if st.button("Generate", type="primary"):
    with st.spinner("Writing the project..."):
        st.session_state["result"] = build_project(
            model, spec, project_name=project_name, mode=mode,
            extract_root=extract_root, guard_rails=guard_rails,
        )

result = st.session_state.get("result")
if result is None:
    st.stop()

for w in result.warnings:
    st.warning(w, icon="⚠️")

zip_bytes = result.to_zip_bytes()
st.download_button(
    f"⬇️ Download {result.project_name}.zip ({len(zip_bytes) / 1024:.0f} KB)",
    zip_bytes,
    file_name=f"{result.project_name}.zip",
    mime="application/zip",
    type="primary",
)

if result.mode == LINKED:
    st.success(
        f"Built as a **living dashboard**. The workbook `{result.workbook_name}` is inside the "
        f"zip under `Data\\` — that is now your master data file.",
        icon="✅",
    )
    st.markdown(
        f"""
**What to do with it**

1. **Extract the zip into `{result.data_folder_hint.rsplit(chr(92) + result.project_name, 1)[0]}`.**
   Not by double-clicking the zip and opening files from inside it — properly, with Extract All.
   The dashboard expects the workbook at `{result.data_folder_hint}`.
2. Open Power BI Desktop → **File → Open → Browse** → pick `{result.project_name}.pbip`.
3. Want an ordinary `.pbix`? Once open: **File → Save As → Power BI files (.pbix)**.

**From now on, to add data:** open the workbook, type into the first empty row inside the
coloured table (blank rows are already waiting — no need to insert any), save, then click
**Refresh** in Power BI.

Extracted somewhere else? **Home → Transform data → Manage Parameters**, set `DataFolder`
to wherever the workbook ended up, then **Close & Apply**.
"""
    )
    if result.sheet_plans:
        st.dataframe(
            pd.DataFrame([
                {
                    "Sheet": p_.sheet_name,
                    "Rows": f"{p_.n_rows:,}",
                    "Blank rows ready": p_.n_spare_rows,
                    "Dropdown columns": len(p_.dropdown_columns),
                    "Headings locked": "yes" if p_.protected else "no",
                }
                for p_ in result.sheet_plans
            ]),
            hide_index=True, use_container_width=True,
        )
else:
    st.markdown(
        f"""
**What to do with it**

1. Extract the whole zip — keep the folders together, the report finds its model by relative path.
2. Open Power BI Desktop → **File → Open → Browse** → pick `{result.project_name}.pbip`.
3. Want a `.pbix`? Once it's open: **File → Save As → Power BI files (.pbix)**.

{"Your data is held inside the file, so there is nothing to point at — but it cannot be refreshed. Generate a living dashboard if you need it to keep up with new data."
 if result.inline_data else
 f"Your data ships as CSV in the `Data` folder, and the model looks for it at `{result.data_folder_hint}`."}
"""
    )

st.divider()
st.caption(
    "Guardrail reminder: Biz-Viz reads your workbook and writes files. It never connects to a "
    "source system and never writes back to one. Relationships and measures are detected "
    "automatically — check them against what you know about the business before this goes to a client."
)
