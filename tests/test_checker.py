import os
import sys

import pandas as pd
import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from generate_data import generate_sites, write_sites_workbook, write_requirements_workbook
from checker import (
    load_sites, load_requirements, find_out_of_order_milestones,
    find_sla_breaches, find_missing_current_milestone_dates,
    find_duplicate_sites, run_all_checks,
)


@pytest.fixture(scope="module")
def workbook_paths(tmp_path_factory):
    rows, problems = generate_sites(n=800, seed=0)
    out_dir = tmp_path_factory.mktemp("data")
    sites_path = str(out_dir / "sites.xlsx")
    reqs_path = str(out_dir / "requirements.xlsx")
    write_sites_workbook(rows, sites_path)
    write_requirements_workbook(reqs_path)
    return sites_path, reqs_path, problems


@pytest.fixture(scope="module")
def dfs(workbook_paths):
    sites_path, reqs_path, problems = workbook_paths
    sites_df = load_sites(sites_path)
    requirements_df = load_requirements(reqs_path)
    return sites_df, requirements_df, problems


def test_load_sites_reads_real_xlsx_file(dfs):
    sites_df, _, _ = dfs
    assert len(sites_df) > 800  # includes injected duplicates
    assert "site_id" in sites_df.columns


def test_load_sites_parses_milestone_dates(dfs):
    sites_df, _, _ = dfs
    assert pd.api.types.is_datetime64_any_dtype(sites_df["datum_Standortakquise"])


def test_load_requirements_reads_real_xlsx_file(dfs):
    _, requirements_df, _ = dfs
    assert len(requirements_df) == 4
    assert requirements_df.iloc[0]["requires"] is None


def test_find_out_of_order_milestones_matches_injected_count(dfs):
    sites_df, requirements_df, problems = dfs
    expected_ids = {site_id for kind, site_id in problems if kind == "out_of_order"}
    result = find_out_of_order_milestones(sites_df, requirements_df)
    found_ids = set(result["site_id"])
    assert expected_ids <= found_ids


def test_find_out_of_order_milestones_flags_correct_pair(dfs):
    sites_df, requirements_df, _ = dfs
    result = find_out_of_order_milestones(sites_df, requirements_df)
    if len(result) > 0:
        row = result.iloc[0]
        assert row["milestone_date"] < row["prerequisite_date"]


def test_find_sla_breaches_all_exceed_max_gap(dfs):
    sites_df, requirements_df, _ = dfs
    result = find_sla_breaches(sites_df, requirements_df)
    assert len(result) > 0
    assert (result["gap_days"] > result["max_allowed_days"]).all()


def test_find_missing_current_milestone_dates_matches_injected_count(dfs):
    sites_df, requirements_df, problems = dfs
    expected_ids = {site_id for kind, site_id in problems if kind == "missing_date"}
    result = find_missing_current_milestone_dates(sites_df)
    found_ids = set(result["site_id"])
    assert expected_ids <= found_ids


def test_find_duplicate_sites_matches_injected_count(dfs):
    sites_df, requirements_df, problems = dfs
    expected_ids = {site_id for kind, site_id in problems if kind == "duplicate"}
    result = find_duplicate_sites(sites_df)
    found_ids = set(result["site_id"])
    assert expected_ids == found_ids
    assert (result["occurrences"] == 2).all()


def test_find_duplicate_sites_no_false_positives_on_unique_ids():
    df = pd.DataFrame({"site_id": ["A", "B", "C"]})
    result = find_duplicate_sites(df)
    assert len(result) == 0


def test_run_all_checks_total_matches_sum_of_parts(dfs):
    sites_df, requirements_df, _ = dfs
    results = run_all_checks(sites_df, requirements_df)
    expected_total = (
        len(results["out_of_order"]) + len(results["sla_breaches"])
        + len(results["missing_dates"]) + len(results["duplicates"])
    )
    assert results["total_issues"] == expected_total


def test_run_all_checks_total_sites_matches_dataframe_length(dfs):
    sites_df, requirements_df, _ = dfs
    results = run_all_checks(sites_df, requirements_df)
    assert results["total_sites"] == len(sites_df)


def test_sla_breach_rate_is_realistic_minority(dfs):
    """
    A genuine, checkable claim about the generator's realism: SLA
    breaches should be a real minority of milestone transitions
    (roughly 8% by design), not dominate the dataset -- otherwise the
    'requirement verification' signal would be meaningless noise.
    """
    sites_df, requirements_df, _ = dfs
    total_transitions = 0
    for _, req in requirements_df.iterrows():
        # `req` comes from .iterrows(), which builds each row as a fresh
        # Series spanning all columns -- pandas can silently upcast a
        # per-column None back to NaN when the row also contains float
        # columns (like max_gap_days). `is None` alone can't be trusted
        # here; pd.isna() correctly catches both representations, same
        # guard used in src/checker.py's own functions.
        if req["requires"] is None or pd.isna(req["requires"]):
            continue
        m_col = f"datum_{req['milestone']}"
        p_col = f"datum_{req['requires']}"
        total_transitions += int((sites_df[m_col].notna() & sites_df[p_col].notna()).sum())

    results = run_all_checks(sites_df, requirements_df)
    breach_rate = len(results["sla_breaches"]) / total_transitions
    assert 0.03 < breach_rate < 0.15
