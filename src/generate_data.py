"""
Synthetic mobile-network rollout data generator.

Disclosure: no real 1&1 rollout data exists or is claimed here. This
generates two realistically-shaped Excel workbooks modeled on the
posting's own description ("Analyse und Vergleich großer
Excel-Datenbestände (z. B. Vertrags-, Termin- und Projektdaten)" and
"Automatisierungen zur Prüfung von Meilensteinen und
Projektanforderungen"):

1. rollout_sites.xlsx -- a site/project tracker: one row per mobile-
   network rollout site, its current milestone stage, planned and
   actual dates, and status.
2. milestone_requirements.xlsx -- the requirement definition sheet:
   for each milestone stage, which prerequisite milestone must already
   be complete and the maximum allowed number of days between them
   (an SLA-style requirement) -- exactly the kind of "Projekt-
   anforderungen" the posting names.

Both are written as real .xlsx files (via openpyxl), not just
DataFrames, so the downstream checker genuinely reads and compares
Excel workbooks, matching the JD's explicit ask.
"""

import random
from datetime import date, timedelta

import openpyxl
from openpyxl import Workbook

MILESTONES = [
    "Standortakquise",
    "Genehmigung",
    "Bau",
    "Inbetriebnahme",
]

# Each milestone's prerequisite and the max allowed gap (days) between
# the prerequisite's completion and this milestone's completion --
# modeling a real rollout SLA/requirement chain.
MILESTONE_REQUIREMENTS = [
    {"milestone": "Standortakquise", "requires": None, "max_gap_days": None},
    {"milestone": "Genehmigung", "requires": "Standortakquise", "max_gap_days": 60},
    {"milestone": "Bau", "requires": "Genehmigung", "max_gap_days": 90},
    {"milestone": "Inbetriebnahme", "requires": "Bau", "max_gap_days": 30},
]

REGIONS = ["NRW", "Bayern", "Baden-Württemberg", "Niedersachsen", "Hessen", "Berlin"]
STATUSES = ["Geplant", "In Bearbeitung", "Abgeschlossen", "Verzögert"]

START_DATE = date(2025, 1, 1)


def _random_date(rng, start, max_offset_days):
    return start + timedelta(days=rng.randint(0, max_offset_days))


def generate_sites(n=800, seed=0):
    """
    Generate n synthetic rollout sites, each with a current milestone
    stage and a realistic completion-date history for the milestones
    already passed. Deliberately injects a known set of data problems
    (a small percentage of sites) so the checker has real issues to
    catch:
      - a milestone completed BEFORE its prerequisite (data-entry error)
      - a gap between milestones exceeding max_gap_days (SLA breach)
      - a missing/blank date for a milestone the site claims to have
        reached
      - a duplicate site ID (two rows for the same site)
    """
    rng = random.Random(seed)
    rows = []

    for i in range(n):
        site_id = f"SITE-{20000 + i}"
        region = rng.choice(REGIONS)

        # How far this site has progressed (0 = only acquired the site,
        # 3 = fully live).
        progress = rng.choices([0, 1, 2, 3], weights=[0.15, 0.25, 0.30, 0.30], k=1)[0]

        dates = {}
        prev_date = _random_date(rng, START_DATE, 300)
        dates["Standortakquise"] = prev_date

        for stage_idx in range(1, progress + 1):
            stage = MILESTONES[stage_idx]
            req = MILESTONE_REQUIREMENTS[stage_idx]
            # Most sites comfortably meet the SLA (gap well under the
            # max); a realistic minority (~8%) overrun it, modeled as a
            # deliberate, disclosed rate rather than an incidental one.
            if rng.random() < 0.08:
                gap = rng.randint(req["max_gap_days"] + 1, req["max_gap_days"] + 30)
            else:
                gap = rng.randint(5, int(req["max_gap_days"] * 0.75))
            stage_date = prev_date + timedelta(days=gap)
            dates[stage] = stage_date
            prev_date = stage_date

        current_stage = MILESTONES[progress]
        status = "Abgeschlossen" if progress == 3 else rng.choice(["Geplant", "In Bearbeitung", "Verzögert"])

        row = {
            "site_id": site_id,
            "region": region,
            "current_milestone": current_stage,
            "status": status,
        }
        for m in MILESTONES:
            row[f"datum_{m}"] = dates.get(m)
        rows.append(row)

    # Inject known data-quality problems into a small, fixed subset so
    # tests can assert exact counts.
    problem_rows = []

    # 1) A milestone completed before its prerequisite (impossible
    #    order) -- pick 10 sites that have reached at least "Bau".
    candidates = [r for r in rows if r["datum_Bau"] is not None]
    for r in rng.sample(candidates, min(10, len(candidates))):
        r["datum_Bau"] = r["datum_Genehmigung"] - timedelta(days=5)
        problem_rows.append(("out_of_order", r["site_id"]))

    # 2) A missing date for a milestone the site claims as its current
    #    stage -- pick 8 sites.
    candidates = [r for r in rows if r["current_milestone"] != "Standortakquise"]
    for r in rng.sample(candidates, min(8, len(candidates))):
        r[f"datum_{r['current_milestone']}"] = None
        problem_rows.append(("missing_date", r["site_id"]))

    # 3) A duplicate site row -- pick 5 sites and append a near-copy.
    dup_candidates = rng.sample(rows, min(5, len(rows)))
    for r in dup_candidates:
        dup = dict(r)
        rows.append(dup)
        problem_rows.append(("duplicate", r["site_id"]))

    rng.shuffle(rows)
    return rows, problem_rows


def write_sites_workbook(rows, path):
    wb = Workbook()
    ws = wb.active
    ws.title = "RolloutSites"

    headers = ["site_id", "region", "current_milestone", "status"] + [f"datum_{m}" for m in MILESTONES]
    ws.append(headers)

    for row in rows:
        ws.append([
            row["site_id"], row["region"], row["current_milestone"], row["status"],
            *[row.get(f"datum_{m}") for m in MILESTONES],
        ])

    wb.save(path)


def write_requirements_workbook(path):
    wb = Workbook()
    ws = wb.active
    ws.title = "MilestoneRequirements"
    ws.append(["milestone", "requires", "max_gap_days"])
    for req in MILESTONE_REQUIREMENTS:
        ws.append([req["milestone"], req["requires"], req["max_gap_days"]])
    wb.save(path)


if __name__ == "__main__":
    rows, problems = generate_sites(n=800, seed=0)
    write_sites_workbook(rows, "data/rollout_sites.xlsx")
    write_requirements_workbook("data/milestone_requirements.xlsx")
    print(f"Wrote {len(rows)} site rows (data/rollout_sites.xlsx)")
    print(f"Wrote milestone requirements (data/milestone_requirements.xlsx)")
    print(f"Injected {len(problems)} known data-quality problems for the checker to find")
