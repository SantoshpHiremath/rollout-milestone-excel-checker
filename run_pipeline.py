"""
End-to-end demo: generate synthetic rollout-site and milestone-
requirement Excel workbooks, run the automated checks, write a real
Excel report, and print a summary.
"""

import sys
sys.path.insert(0, "src")

from generate_data import generate_sites, write_sites_workbook, write_requirements_workbook
from checker import load_sites, load_requirements, run_all_checks
from report import write_report


def main():
    print("=" * 70)
    print("MOBILFUNK-ROLLOUT MEILENSTEIN-PRÜFUNG (Excel-Automatisierung)")
    print("=" * 70)

    rows, problems = generate_sites(n=800, seed=0)
    write_sites_workbook(rows, "data/rollout_sites.xlsx")
    write_requirements_workbook("data/milestone_requirements.xlsx")
    print(f"\n{len(rows)} Standort-Zeilen generiert (data/rollout_sites.xlsx)")
    print(f"{len(problems)} bekannte Datenprobleme absichtlich eingefügt (zum Testen des Checkers)")

    sites_df = load_sites("data/rollout_sites.xlsx")
    requirements_df = load_requirements("data/milestone_requirements.xlsx")

    results = run_all_checks(sites_df, requirements_df)

    print(f"\nGeprüfte Standorte: {results['total_sites']}")
    print(f"Gefundene Auffälligkeiten gesamt: {results['total_issues']}")
    print(f"  - Falsche Meilenstein-Reihenfolge: {len(results['out_of_order'])}")
    print(f"  - SLA-/Anforderungsverletzungen: {len(results['sla_breaches'])}")
    print(f"  - Fehlende Meilenstein-Daten: {len(results['missing_dates'])}")
    print(f"  - Doppelte Standort-Einträge: {len(results['duplicates'])}")

    out_path = write_report(results, "output/milestone_check_report.xlsx")
    print(f"\nExcel-Report geschrieben: {out_path}")

    print("\n" + "=" * 70)
    print("Hinweis: Die Rollout-Daten sind synthetisch (siehe README), da")
    print("keine echten 1&1-Daten in dieser Umgebung verfügbar sind. Die")
    print("Meilenstein-Struktur und Anforderungslogik orientieren sich")
    print("direkt an der Aufgabenbeschreibung dieser Ausschreibung.")
    print("=" * 70)


if __name__ == "__main__":
    main()
