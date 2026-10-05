# rollout-milestone-excel-checker

An Excel-dataset comparison and milestone/requirement-verification
automation tool. It reads large Excel datasets, checks them against
milestone and project requirements, and writes the findings into a
formatted Excel report, the kind of repeatable check that replaces
manual spreadsheet comparison in a rollout project.

## What it does

- `src/generate_data.py` — generates two real Excel workbooks modeling
  a mobile-network rollout: `rollout_sites.xlsx` (one row per rollout
  site, its current milestone, and the dates it reached each milestone
  stage) and `milestone_requirements.xlsx` (the requirement chain: which
  milestone requires which prior milestone, and the maximum allowed gap
  in days between them — a real SLA-style requirement definition).
  Deliberately injects a small, known set of data-quality
  problems (out-of-order milestone dates, SLA breaches, missing dates,
  duplicate site rows) so the checker has real issues to find.
- `src/checker.py` — the real automation: reads both Excel workbooks
  with pandas/openpyxl and runs four independent checks — impossible
  milestone ordering, requirement/SLA breaches, missing dates for a
  site's claimed current stage, and duplicate site entries.
- `src/report.py` — writes the results into a real, formatted Excel
  report (`output/milestone_check_report.xlsx`) with a summary sheet
  plus one sheet per issue type — the reporting side of the check.
- `run_pipeline.py` — runs the full flow end to end and prints a real
  summary (see "Sample output" below, copied directly from an actual
  run).
- `tests/` — 22 tests, all passing, covering the data generator, all
  four checker functions (each verified against the generator's own
  known-injected problems, not just "runs without crashing"), and the
  Excel report writer.

## Scope

**The rollout data is synthetic.** No real mobile-network rollout
dataset was used. The milestone structure (Standortakquise →
Genehmigung → Bau → Inbetriebnahme) and the requirement-chain concept
(each milestone has a prerequisite and a maximum allowed gap) are a
modeling choice, not a description of any real rollout process or SLAs.

**The data-quality problems are deliberately injected and their exact
counts are known**, so the test suite can assert precise outcomes
(`test_find_out_of_order_milestones_matches_injected_count`, etc.)
rather than just checking "some issues were found." This is a planted-problems pattern, as in the SAP+Excel controlling-report project's deliberately malformed SAP export.

**Both input files and the output report are real `.xlsx` workbooks**,
read and written with `openpyxl`/`pandas`, not CSVs relabeled as Excel.
I checked directly: `pd.read_excel(..., engine="openpyxl")` and
`openpyxl.load_workbook(...)` both parse the files, and the
report-writer tests open the actual output file and read real cell
values back out of it.

**SLA breach rate is calibrated to be a realistic minority** (~8% of
milestone transitions), not the majority. An earlier version of the
generator made breaches occur far too often (over 100% of some
milestone transitions when measured against the wrong denominator),
which would have made the "requirement verification" signal
meaningless. I caught it by comparing the checker's counts against the
actual number of eligible milestone transitions (not just site count)
and fixed it, with `test_sla_breach_rate_is_realistic_minority` guarding
against a regression.

## Sample output (from an actual run of `run_pipeline.py`)

```
805 Standort-Zeilen generiert (data/rollout_sites.xlsx)
23 bekannte Datenprobleme absichtlich eingefügt (zum Testen des Checkers)

Geprüfte Standorte: 805
Gefundene Auffälligkeiten gesamt: 134
  - Falsche Meilenstein-Reihenfolge: 10
  - SLA-/Anforderungsverletzungen: 111
  - Fehlende Meilenstein-Daten: 8
  - Doppelte Standort-Einträge: 5

Excel-Report geschrieben: output/milestone_check_report.xlsx
```

The checker recovers all 23 deliberately injected problems (10
out-of-order, 8 missing-date, 5 duplicate) plus 111 SLA breaches that
emerge naturally from the generator's randomized but realistically-
calibrated milestone gaps — evidence the checks generalize beyond just
the hand-planted cases.

## Testing

- `python3 -m pytest tests/ -v` — 22/22 tests pass.
- `python3 run_pipeline.py` — runs end to end against real `.xlsx`
  files; the "Sample output" section above is copied directly from
  this run's actual stdout.
- Each of the four checker functions is tested against the generator's
  own known-injected problem set (exact site IDs), not just a
  plausible-looking count.
- The Excel report writer is tested by re-opening the actual output
  file with `openpyxl.load_workbook` and reading real cell values back
  out, confirming the report's summary counts match the checker's
  results exactly.

## Running it yourself

```bash
pip install -r requirements.txt
python3 -m pytest tests/ -v
python3 run_pipeline.py
```
