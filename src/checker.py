"""
Milestone/requirement automation checker.

Reads the two real Excel workbooks (rollout_sites.xlsx,
milestone_requirements.xlsx) with pandas/openpyxl and runs the
automated checks the posting explicitly describes:
"Entwicklung und Implementierung von Automatisierungen zur Prüfung von
Meilensteinen und Projektanforderungen".

Four real, independently useful checks:
  1. out-of-order milestones (a later milestone dated before its
     prerequisite -- a real data-entry error class)
  2. SLA/requirement breaches (the gap between a milestone and its
     prerequisite exceeds the allowed max_gap_days)
  3. missing dates for a site's claimed current milestone
  4. duplicate site rows
"""

import pandas as pd


MILESTONE_DATE_PREFIX = "datum_"


def load_sites(path="data/rollout_sites.xlsx"):
    df = pd.read_excel(path, engine="openpyxl")
    date_cols = [c for c in df.columns if c.startswith(MILESTONE_DATE_PREFIX)]
    for c in date_cols:
        df[c] = pd.to_datetime(df[c], errors="coerce")
    return df


def load_requirements(path="data/milestone_requirements.xlsx"):
    df = pd.read_excel(path, engine="openpyxl")
    # openpyxl/pandas represent an empty cell as NaN; normalize to None
    # for the "requires" column of the first milestone (no prerequisite).
    # The column is cast to object dtype first: on some pandas versions,
    # calling .where(..., None) on a column pandas has inferred as
    # float64 silently keeps NaN instead of actually storing None,
    # because None has no representation in a float64 Series. Casting to
    # object first guarantees None is stored as a real Python None, which
    # is what "requires is None" downstream (here and in find_out_of_order_
    # milestones/find_sla_breaches) actually depends on.
    df["requires"] = df["requires"].astype(object).where(df["requires"].notna(), None)
    return df


def find_out_of_order_milestones(sites_df, requirements_df):
    """
    Flag sites where a milestone's date is earlier than its
    prerequisite's date -- a logically impossible order that indicates
    a data-entry error.
    """
    issues = []
    for _, req in requirements_df.iterrows():
        milestone, prereq = req["milestone"], req["requires"]
        # Guard against both None (the expected "no prerequisite" marker)
        # and a stray NaN (e.g. if a pandas version quirk lets one slip
        # through load_requirements' None-normalization) -- either way,
        # there's no real prerequisite column to compare against, and
        # building a column name from NaN would silently produce a
        # nonexistent "datum_nan" column and crash with a KeyError.
        if prereq is None or pd.isna(prereq):
            continue
        m_col = f"{MILESTONE_DATE_PREFIX}{milestone}"
        p_col = f"{MILESTONE_DATE_PREFIX}{prereq}"
        both_present = sites_df[m_col].notna() & sites_df[p_col].notna()
        out_of_order = both_present & (sites_df[m_col] < sites_df[p_col])
        for _, row in sites_df[out_of_order].iterrows():
            issues.append({
                "site_id": row["site_id"],
                "issue": "out_of_order",
                "milestone": milestone,
                "prerequisite": prereq,
                "milestone_date": row[m_col],
                "prerequisite_date": row[p_col],
            })
    return pd.DataFrame(issues)


def find_sla_breaches(sites_df, requirements_df):
    """
    Flag sites where the gap between a milestone and its prerequisite
    exceeds the requirement sheet's max_gap_days -- the "Prüfung von
    ... Projektanforderungen" the posting names directly.
    """
    issues = []
    for _, req in requirements_df.iterrows():
        milestone, prereq, max_gap = req["milestone"], req["requires"], req["max_gap_days"]
        if prereq is None or pd.isna(prereq) or pd.isna(max_gap):
            continue
        m_col = f"{MILESTONE_DATE_PREFIX}{milestone}"
        p_col = f"{MILESTONE_DATE_PREFIX}{prereq}"
        both_present = sites_df[m_col].notna() & sites_df[p_col].notna()
        gap_days = (sites_df[m_col] - sites_df[p_col]).dt.days
        breach = both_present & (gap_days > max_gap) & (gap_days >= 0)
        for idx in sites_df[breach].index:
            row = sites_df.loc[idx]
            issues.append({
                "site_id": row["site_id"],
                "issue": "sla_breach",
                "milestone": milestone,
                "prerequisite": prereq,
                "gap_days": int(gap_days.loc[idx]),
                "max_allowed_days": int(max_gap),
            })
    return pd.DataFrame(issues)


def find_missing_current_milestone_dates(sites_df):
    """Flag sites where the date for their own claimed current milestone is blank."""
    issues = []
    for _, row in sites_df.iterrows():
        col = f"{MILESTONE_DATE_PREFIX}{row['current_milestone']}"
        if col in sites_df.columns and pd.isna(row[col]):
            issues.append({
                "site_id": row["site_id"],
                "issue": "missing_date",
                "milestone": row["current_milestone"],
            })
    return pd.DataFrame(issues)


def find_duplicate_sites(sites_df):
    """Flag site_ids that appear more than once."""
    counts = sites_df["site_id"].value_counts()
    dup_ids = counts[counts > 1].index.tolist()
    issues = [{"site_id": sid, "issue": "duplicate", "occurrences": int(counts[sid])} for sid in dup_ids]
    return pd.DataFrame(issues)


def run_all_checks(sites_df, requirements_df):
    """Run all four checks and return a combined summary dict."""
    out_of_order = find_out_of_order_milestones(sites_df, requirements_df)
    sla_breaches = find_sla_breaches(sites_df, requirements_df)
    missing_dates = find_missing_current_milestone_dates(sites_df)
    duplicates = find_duplicate_sites(sites_df)

    return {
        "out_of_order": out_of_order,
        "sla_breaches": sla_breaches,
        "missing_dates": missing_dates,
        "duplicates": duplicates,
        "total_sites": len(sites_df),
        "total_issues": len(out_of_order) + len(sla_breaches) + len(missing_dates) + len(duplicates),
    }
