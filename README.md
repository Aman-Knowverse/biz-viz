# Biz-Viz

Upload a client's Excel workbook, get back a Power BI project that opens in Power BI Desktop
with the data model and a starter dashboard already built.

## Why `.pbip` and not `.pbix`

A `.pbix` is an undocumented binary container. Nothing outside Power BI itself can reliably
write one, and any library claiming to is either reading it or faking it.

A `.pbip` is Microsoft's **project** format: the same report stored as plain text — TMDL for
the semantic model, JSON per visual for the report. Power BI Desktop opens it natively, and
it is becoming the default report format. Once it's open, `File → Save As → .pbix` gives you
the binary if you need one.

So Biz-Viz generates text files, which is a problem that can actually be solved correctly,
rather than a binary format that can only be approximated.

## Two things it can produce

**Living dashboard** (default). The zip contains a cleaned Excel workbook *and* the
project, with the model reading from that workbook through Power Query. The user
keeps filling the workbook in; Refresh moves the dashboard. This is the one that
survives contact with a real engagement.

**Snapshot.** The data is compressed inside the model itself. Nothing to point at,
no paths to fix, opens anywhere — and can never be refreshed. The right thing to
email a client.

```
Excel upload
   ↓  PROFILE      column types, cardinality, date detection, measure candidates
   ↓  CLEAN        skip title blocks, drop blank + subtotal rows, fill merged
                   label cells, convert "₹ 1,23,456.00" text into real numbers
   ↓  MODEL        detect relationships, prune to a valid star schema,
                   write DAX measures
   ↓  DESIGN       rules or Claude decide the KPIs, pages and business names
   ↓  WORKBOOK     one sheet per table as a named Excel Table, with spare rows,
                   dropdowns, locked headings and a READ ME sheet  (linked only)
   ↓  WRITE        model.tmdl reading the workbook via Power Query,
                   + definition/pages/**/visual.json
   ↓  PACKAGE      zip the project, the workbook and START_HERE.txt
```

## Layout

| File | Does |
|---|---|
| `app.py` | Streamlit UI — upload, review, adjust roles, generate, download |
| `bizviz/profiling.py` | Read and clean sheets; infer what each column *means* |
| `bizviz/modeling.py` | Relationships, calendar table, DAX measures |
| `bizviz/design.py` | Decide the dashboard — rule-based and Claude designers |
| `bizviz/excel_writer.py` | Write the master workbook the user keeps filling in |
| `bizviz/tmdl.py` | Write the semantic model (TMDL + embedded Power Query) |
| `bizviz/pbir.py` | Write the report (PBIR visual/page/report JSON) |
| `bizviz/build.py` | Assemble the project folder, README and zip |

## Run it

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Four design decisions worth knowing about

**The reshaping happens in Power Query, not in the file.** If a sheet was a crosstab with
months running across, the workbook *keeps* that layout and the unpivot is written as visible
query steps. Baking tidy long-format data into the workbook instead would be correct for Power
BI and miserable for the human who has to keep filling it in — they would be typing one row per
month forever. This way they carry on filling across, and Power BI rearranges on every refresh.

**The calendar is built in Power Query too.** A calendar generated once at build time silently
stops covering the data the first time someone enters a date in a new year, and the trend chart
just drops those rows. It is rebuilt from the actual min and max on every refresh. A DAX
calculated table would also work, but calculated-column TMDL syntax could not be verified
against anything authoritative — and an unverifiable guess that breaks the entire model is a
poor trade for something M does perfectly well.

**Power Query has no relative paths.** A real Microsoft limitation, not an oversight — still an
open feature request. So the folder is stored as a *parameter* (`DataFolder`), which means a
user who extracted the zip somewhere unexpected fixes it by typing a folder into Power BI's
Manage Parameters dialog rather than editing a query. The app asks up front where they intend
to extract to, so most of the time there is nothing to fix.

In snapshot mode there is no path at all: the data is deflate-compressed into the Power Query
expression itself, the same mechanism Power BI's own "Enter Data" uses.

**Value fields bind to columns, not to model measures.** This one was learned the hard
way. The generator wrote 18 measures into the model, the report referenced them by name,
every validator passed — and Power BI Desktop then quietly declined to create the measures,
leaving every measure-bound visual showing `Missing_References` on an otherwise healthy
model. Columns, relationships, queries and the calendar were all fine.

So a value field now binds to the column underneath with an explicit `Aggregation` — the
same expression Power BI writes when a person drags a numeric field onto a chart, and
therefore something that cannot be rejected for any reason the columns themselves survive.
The measures are still written into the model; they are simply no longer load-bearing. The
cost is cosmetic: a card reads "Sum of Revenue" rather than "Total Revenue".

**Relationships are pruned to a spanning forest.** Name-matching finds far more joins than
Power BI will accept — a fact joined to a dimension on both its ID and its name, or a loop
between three tables. Power BI allows exactly one active filter path between any two tables,
so the candidates are sorted by trust and accepted only when they connect tables not already
linked. Everything discarded is reported in the UI rather than hidden.

## Where Claude fits

In Flow Finder, Claude was a **narrator** — it read what pandas discovered and wrote the
diagnosis. Here it is a **designer**: it reads the model's *structure* (table names, column
names, types, distinct counts — never the rows) and returns a plan for which measures are
headline KPIs, how the pages break up, and what to call things in client language.

Every field it names is checked against the real model before use. Anything that doesn't
exist is discarded and the rule-based choice stands. Claude can change the plan; it never
writes a byte of the output.

The whole thing also works with no API key at all — the rules-based designer is the default.

## Verifying changes

```bash
python run_test.py /tmp/test_clean.xlsx /tmp/out    # end-to-end, prints what was detected
python verify_bindings.py                           # every field a visual asks for resolves to
                                                    # a real column, no measure dependencies,
                                                    # aggregation naming conventions correct
python verify_linked.py /tmp/test_clean.xlsx        # replay the Power Query steps against the
                                                    # real generated workbook and prove they
                                                    # reproduce the tidy data exactly
python verify_model.py /tmp/test_clean.xlsx         # snapshot mode: decode embedded data
python validate_schemas.py <path>.Report            # validate against Microsoft's JSON schemas
python test_edges.py                                # nasty names, dirty numbers, mixed dates...
python verify_all.py                                # everything, over every test workbook
```

`validate_schemas.py` needs a clone of Microsoft's schema repo at `/tmp/json-schemas`:

```bash
git clone --depth 1 https://github.com/microsoft/json-schemas.git /tmp/json-schemas
```

Microsoft also publishes a PBIR validator, which is worth running:

```bash
npm install -g @microsoft/powerbi-report-authoring-cli
powerbi-report-author validate "<Project>/<Project>.Report"
```

## Status

The generated report passes Microsoft's own PBIR validator with zero errors, and every
generated JSON file validates against Microsoft's published schemas. The embedded data
round-trips exactly and is type-safe against its declared Power Query conversions. For the
linked mode, the generated workbook is reopened and every Power Query step replayed in pandas,
proving the reshaping reproduces the tidy data the model expects — cell for cell, across ten
test workbooks including crosstabs, dirty numbers and hostile column names.

**Confirmed in Power BI Desktop (Sept 2026):** the project opens, the semantic model loads, the
Power Query reads the generated workbook, the relationships resolve without ambiguity errors and
the generated calendar builds. Column-bound visuals render.

**The one failure found in real use**, and what it cost: measures defined in the model file were
silently dropped by Power BI Desktop, so every measure-bound visual failed. Root cause on
Microsoft's side is still unknown — the TMDL matched their documented example exactly. The
generator routes around it by binding values to columns instead, and `verify_bindings.py` now
fails the build if anything reintroduces a dependency on a model measure.

The lesson worth keeping: passing Microsoft's own schema and PBIR validators proves a file is
*well-formed*, not that Power BI will *honour* it. Only opening it does that.
