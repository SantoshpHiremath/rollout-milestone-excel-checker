# rollout-milestone-excel-checker

A real Excel-dataset comparison and milestone/requirement-verification
automation tool — built to close the one specific, closable gap
identified against 1&1's "Werkstudent AI & Data Automation - Mobilfunk
Rollout" posting: automated checking of large Excel datasets against
milestone/project requirements, exactly as the posting describes
("Analyse und Vergleich großer Excel-Datenbestände... Automatisierungen
zur Prüfung von Meilensteinen und Projektanforderungen").

## What this is, precisely

- `src/generate_data.py` — generates two real Excel workbooks modeling
  a mobile-network rollout: `rollout_sites.xlsx` (one row per rollout
  site, its current milestone, and the dates it reached each milestone
  stage) and `milestone_requirements.xlsx` (the requirement chain: which
  milestone requires which prior milestone, and the maximum allowed gap
  in days between them — a real SLA-style requirement definition).
  Deliberately injects a small, known, disclosed set of data-quality
  problems (out-of-order milestone dates, SLA breaches, missing dates,
  duplicate site rows) so the checker has real issues to find.
- `src/checker.py` — the real automation: reads both Excel workbooks
  with pandas/openpyxl and runs four independent checks — impossible
  milestone ordering, requirement/SLA breaches, missing dates for a
  site's claimed current stage, and duplicate site entries.
- `src/report.py` — writes the results into a real, formatted Excel
  report (`output/milestone_check_report.xlsx`) with a summary sheet
  plus one sheet per issue type — the "Erstellung ... von Reports und
  Kennzahlen" ask from the posting.
- `run_pipeline.py` — runs the full flow end to end and prints a real
  summary (see "Sample output" below, copied directly from an actual
  run).
- `tests/` — 22 tests, all passing, covering the data generator, all
  four checker functions (each verified against the generator's own
  known-injected problems, not just "runs without crashing"), and the
  Excel report writer.

## Honest disclosure — what's real, what's substituted, and why

**The rollout data is synthetic, not real 1&1 data.** No real mobile-
network rollout dataset exists or was claimed here — 1&1's actual site,
contract, and project data is obviously not accessible in this
environment. The milestone structure (Standortakquise → Genehmigung →
Bau → Inbetriebnahme) and the requirement-chain concept (each milestone
has a prerequisite and a maximum allowed gap) are a reasonable, disclosed
modeling choice built directly from the posting's own wording, not a
claim about 1&1's actual rollout process or SLAs.

**The data-quality problems are deliberately injected and their exact
counts are known**, specifically so the test suite can assert precise
outcomes (`test_find_out_of_order_milestones_matches_injected_count`,
etc.) rather than just checking "some issues were found." This mirrors
the same disclosed-synthetic-data-with-planted-problems pattern used in
other projects in this portfolio (e.g. the SAP+Excel controlling-report
project's deliberately malformed SAP export).

**Both input files and the output report are real `.xlsx` workbooks**,
read and written with `openpyxl`/`pandas`, not CSVs relabeled as Excel —
this was checked directly: `pd.read_excel(..., engine="openpyxl")` and
`openpyxl.load_workbook(...)` both genuinely parse the files, and the
report-writer tests open the actual output file and read real cell
values back out of it.

**SLA breach rate is deliberately calibrated to be a realistic
minority** (~8% of milestone transitions), not the majority. An earlier
version of the generator made breaches occur far too often (over 100%
of some milestone transitions when measured against the wrong
denominator), which would have made the "requirement verification"
signal meaningless — this was caught by comparing the checker's counts
against the actual number of eligible milestone transitions (not just
site count) and fixed, with `test_sla_breach_rate_is_realistic_minority`
guarding against a regression.

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

## Verification performed

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
