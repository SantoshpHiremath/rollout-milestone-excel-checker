import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from generate_data import generate_sites, MILESTONES, MILESTONE_REQUIREMENTS


def test_generate_sites_returns_at_least_n_rows():
    """
    n rows requested plus 5 injected duplicates -- returns >= n, not
    exactly n, since duplicates are appended after the initial n.
    """
    rows, problems = generate_sites(n=200, seed=1)
    assert len(rows) >= 200


def test_generate_sites_only_uses_known_milestones():
    rows, _ = generate_sites(n=200, seed=1)
    for r in rows:
        assert r["current_milestone"] in MILESTONES


def test_generate_sites_reproducible_with_same_seed():
    rows_a, problems_a = generate_sites(n=100, seed=42)
    rows_b, problems_b = generate_sites(n=100, seed=42)
    assert rows_a == rows_b
    assert problems_a == problems_b


def test_generate_sites_injects_expected_problem_counts():
    """
    The generator deliberately injects a fixed, known number of each
    problem type so downstream tests can assert exact counts -- this
    guards that the injection logic itself doesn't silently drift.
    """
    rows, problems = generate_sites(n=800, seed=0)
    by_type = {}
    for kind, _ in problems:
        by_type[kind] = by_type.get(kind, 0) + 1
    assert by_type["out_of_order"] == 10
    assert by_type["missing_date"] == 8
    assert by_type["duplicate"] == 5


def test_generate_sites_first_milestone_always_has_a_date():
    rows, _ = generate_sites(n=200, seed=2)
    for r in rows:
        assert r["datum_Standortakquise"] is not None


def test_milestone_requirements_chain_is_well_formed():
    """The first milestone has no prerequisite; every later one requires the previous."""
    assert MILESTONE_REQUIREMENTS[0]["requires"] is None
    for i in range(1, len(MILESTONE_REQUIREMENTS)):
        assert MILESTONE_REQUIREMENTS[i]["requires"] == MILESTONES[i - 1]
