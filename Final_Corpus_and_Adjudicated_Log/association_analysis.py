#!/usr/bin/env python3
"""
Calculate AIFTax cross-dimensional associations.

Analyses:
  1. Failure Category × Risk Assessment
  2. Failure Category × Recovery Complexity

Outputs:
  - contingency tables
  - Pearson chi-square statistics
  - asymptotic p-values
  - fixed-margin Monte Carlo p-values
  - standard and bias-corrected Cramér's V
  - upper-level risk/recovery summaries

The standard Cramér's V values are the values intended for the paper.

Install:
    pip install pandas scipy openpyxl

Run:
   python association_analysis.py \
     --input "AIFTax_Final_Adjudicated_Dataset.xlsx" \
     --output-dir "Cross_Dimensional_Analysis"
"""

from __future__ import annotations

import argparse
import json
import math
from pathlib import Path
from typing import Any

import numpy as np
import pandas as pd
from scipy.stats import chi2_contingency, random_table


REQUIRED_COLUMNS = [
    "Failure Category",
    "Risk Assessment",
    "Recovery Complexity",
]

FAILURE_ORDER = [
    "Operational",
    "Distributional",
    "Adversarial",
    "Mixed/Hybrid",
]

RISK_ORDER = ["Critical", "Severe", "High", "Moderate", "Low"]
RECOVERY_ORDER = ["Very High", "Severe", "High", "Moderate", "Low"]

RISK_HIGH_OR_WORSE = {"Critical", "Severe", "High"}
RECOVERY_HIGH_OR_WORSE = {"Very High", "Severe", "High"}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Calculate Cramér's V for the AIFTax principal variables."
    )
    parser.add_argument(
        "--input",
        required=True,
        type=Path,
        help="Final adjudicated .xlsx, .xlsm, or .csv dataset.",
    )
    parser.add_argument(
        "--sheet-name",
        default="Final Corpus",
        help='Excel sheet name (default: "Final Corpus").',
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=Path("association_results"),
        help='Output directory (default: "association_results").',
    )
    parser.add_argument(
        "--simulations",
        type=int,
        default=200_000,
        help="Fixed-margin Monte Carlo samples (default: 200000).",
    )
    parser.add_argument(
        "--seed",
        type=int,
        default=20260731,
        help=(
            "Random seed for the risk test; the recovery test uses seed + 1 "
            "(default: 20260731)."
        ),
        # help="Random seed (default: 20260731).",
    )
    parser.add_argument(
        "--batch-size",
        type=int,
        default=20_000,
        help="Monte Carlo batch size (default: 20000).",
    )
    return parser.parse_args()


def clean_name(value: Any) -> str:
    return " ".join(str(value).strip().split())


def find_excel_header(path: Path, sheet_name: str) -> int:
    preview = pd.read_excel(
        path,
        sheet_name=sheet_name,
        header=None,
        nrows=15,
        dtype=object,
    )
    required = set(REQUIRED_COLUMNS)

    for row_index, row in preview.iterrows():
        values = {
            clean_name(value)
            for value in row.tolist()
            if pd.notna(value)
        }
        if required.issubset(values):
            return int(row_index)

    raise ValueError(
        "Could not locate an Excel header row containing: "
        + ", ".join(REQUIRED_COLUMNS)
    )


def load_dataset(path: Path, sheet_name: str) -> pd.DataFrame:
    if not path.exists():
        raise FileNotFoundError(f"Input file not found: {path}")

    suffix = path.suffix.lower()
    if suffix in {".xlsx", ".xlsm"}:
        header_row = find_excel_header(path, sheet_name)
        frame = pd.read_excel(
            path,
            sheet_name=sheet_name,
            header=header_row,
            dtype=object,
        )
    elif suffix == ".csv":
        frame = pd.read_csv(path, dtype=object)
    else:
        raise ValueError("Input must be .xlsx, .xlsm, or .csv.")

    frame.columns = [clean_name(column) for column in frame.columns]

    missing = sorted(set(REQUIRED_COLUMNS).difference(frame.columns))
    if missing:
        raise ValueError("Missing required columns: " + ", ".join(missing))

    frame = frame[REQUIRED_COLUMNS].copy().dropna(how="all")
    for column in REQUIRED_COLUMNS:
        frame[column] = frame[column].map(
            lambda value: clean_name(value) if pd.notna(value) else np.nan
        )

    missing_counts = frame.isna().sum()
    if int(missing_counts.sum()) > 0:
        detail = ", ".join(
            f"{column}={int(count)}"
            for column, count in missing_counts.items()
            if count
        )
        raise ValueError("Missing analytical values: " + detail)

    return frame.reset_index(drop=True)


def validate_labels(
    frame: pd.DataFrame,
    column: str,
    allowed: list[str],
) -> None:
    unexpected = sorted(set(frame[column].unique()).difference(allowed))
    if unexpected:
        raise ValueError(
            f"Unexpected labels in {column}: {unexpected}. "
            f"Allowed labels: {allowed}"
        )


def make_table(
    frame: pd.DataFrame,
    outcome: str,
    outcome_order: list[str],
) -> pd.DataFrame:
    table = pd.crosstab(frame["Failure Category"], frame[outcome])
    table = table.reindex(
        index=FAILURE_ORDER,
        columns=outcome_order,
        fill_value=0,
    )
    table = table.loc[table.sum(axis=1) > 0, table.sum(axis=0) > 0]
    return table.astype(int)


def association_statistics(table: pd.DataFrame) -> dict[str, float | int]:
    observed = table.to_numpy(dtype=float)
    n = int(observed.sum())
    rows, columns = observed.shape

    if n == 0 or min(rows, columns) < 2:
        raise ValueError("Cramér's V requires a non-empty table of at least 2×2.")

    chi2, asymptotic_p, dof, expected = chi2_contingency(
        observed,
        correction=False,
    )

    # Standard Cramér's V:
    # V = sqrt(chi-square / (n * min(r - 1, c - 1)))
    standard_v = math.sqrt(
        chi2 / (n * min(rows - 1, columns - 1))
    )

    # Bergsma/Wicher finite-sample bias correction.
    phi2 = chi2 / n
    phi2_corrected = max(
        0.0,
        phi2 - ((columns - 1) * (rows - 1)) / (n - 1),
    )
    rows_corrected = rows - ((rows - 1) ** 2) / (n - 1)
    columns_corrected = columns - ((columns - 1) ** 2) / (n - 1)
    denominator = min(rows_corrected - 1, columns_corrected - 1)
    corrected_v = (
        math.sqrt(phi2_corrected / denominator)
        if denominator > 0
        else float("nan")
    )

    below_5 = int(np.count_nonzero(expected < 5))
    total_cells = int(expected.size)

    return {
        "n": n,
        "chi_square": float(chi2),
        "degrees_freedom": int(dof),
        "asymptotic_p": float(asymptotic_p),
        "cramers_v": float(standard_v),
        "bias_corrected_cramers_v": float(corrected_v),
        "expected_cells_below_5": below_5,
        "total_cells": total_cells,
        "percent_expected_cells_below_5": 100.0 * below_5 / total_cells,
        "minimum_expected_count": float(expected.min()),
    }


def monte_carlo_p(
    table: pd.DataFrame,
    simulations: int,
    seed: int,
    batch_size: int,
) -> float:
    """
    Fixed-margin Monte Carlo test of independence.

    Random tables are generated under the null hypothesis while preserving
    the observed row and column totals. Pearson's chi-square is the test
    statistic.
    """
    if simulations < 1 or batch_size < 1:
        raise ValueError("Simulation count and batch size must be positive.")

    observed = table.to_numpy(dtype=np.int64)
    row_totals = observed.sum(axis=1)
    column_totals = observed.sum(axis=0)
    n = int(observed.sum())
    expected = np.outer(row_totals, column_totals) / n

    observed_chi2 = float(
        np.sum((observed - expected) ** 2 / expected)
    )

    rng = np.random.default_rng(seed)
    extreme = 0
    remaining = simulations

    while remaining > 0:
        size = min(batch_size, remaining)
        sampled = random_table.rvs(
            row_totals,
            column_totals,
            size=size,
            random_state=rng,
        )
        sampled_chi2 = np.sum(
            (sampled - expected) ** 2 / expected,
            axis=(1, 2),
        )
        extreme += int(
            np.count_nonzero(sampled_chi2 >= observed_chi2 - 1e-12)
        )
        remaining -= size

    return (extreme + 1) / (simulations + 1)


def analyze(
    label: str,
    table: pd.DataFrame,
    simulations: int,
    seed: int,
    batch_size: int,
) -> dict[str, Any]:
    result = association_statistics(table)
    result["analysis"] = label
    result["monte_carlo_p"] = monte_carlo_p(
        table,
        simulations=simulations,
        seed=seed,
        batch_size=batch_size,
    )
    result["monte_carlo_simulations"] = simulations
    result["seed"] = seed
    return result


def upper_level_summary(frame: pd.DataFrame) -> pd.DataFrame:
    rows: list[dict[str, Any]] = []

    for category in FAILURE_ORDER:
        subset = frame[frame["Failure Category"] == category]
        if subset.empty:
            continue

        n = len(subset)
        risk_n = int(
            subset["Risk Assessment"].isin(RISK_HIGH_OR_WORSE).sum()
        )
        recovery_n = int(
            subset["Recovery Complexity"]
            .isin(RECOVERY_HIGH_OR_WORSE)
            .sum()
        )

        rows.append(
            {
                "Failure Category": category,
                "n": n,
                "High-or-worse Risk Count": risk_n,
                "High-or-worse Risk Percent": 100.0 * risk_n / n,
                "High-or-worse Recovery Count": recovery_n,
                "High-or-worse Recovery Percent": 100.0 * recovery_n / n,
            }
        )

    return pd.DataFrame(rows)


def main() -> int:
    args = parse_args()

    frame = load_dataset(args.input, args.sheet_name)
    validate_labels(frame, "Failure Category", FAILURE_ORDER)
    validate_labels(frame, "Risk Assessment", RISK_ORDER)
    validate_labels(frame, "Recovery Complexity", RECOVERY_ORDER)

    risk_table = make_table(frame, "Risk Assessment", RISK_ORDER)
    recovery_table = make_table(
        frame,
        "Recovery Complexity",
        RECOVERY_ORDER,
    )

    risk_result = analyze(
        "Failure Category × Risk Assessment",
        risk_table,
        args.simulations,
        args.seed,
        args.batch_size,
    )
    recovery_result = analyze(
        "Failure Category × Recovery Complexity",
        recovery_table,
        args.simulations,
        args.seed + 1,
        args.batch_size,
    )

    args.output_dir.mkdir(parents=True, exist_ok=True)

    risk_table.to_csv(
        args.output_dir / "failure_category_by_risk.csv"
    )
    recovery_table.to_csv(
        args.output_dir / "failure_category_by_recovery.csv"
    )
    upper_level_summary(frame).to_csv(
        args.output_dir / "upper_level_summary.csv",
        index=False,
    )

    results = pd.DataFrame([risk_result, recovery_result])
    results.to_csv(
        args.output_dir / "association_results.csv",
        index=False,
    )

    metadata = {
        "input": str(args.input),
        "sheet_name": args.sheet_name,
        "record_count": int(len(frame)),
        "results": [risk_result, recovery_result],
    }
    (args.output_dir / "association_results.json").write_text(
        json.dumps(metadata, indent=2),
        encoding="utf-8",
    )

    print(f"Loaded {len(frame)} records from {args.input}\n")
    for result in (risk_result, recovery_result):
        print(result["analysis"])
        print("-" * len(result["analysis"]))
        print(f"Pearson chi-square: {result['chi_square']:.6f}")
        print(f"Degrees of freedom: {result['degrees_freedom']}")
        print(f"Asymptotic p-value: {result['asymptotic_p']:.6f}")
        print(
            "Fixed-margin Monte Carlo p-value: "
            f"{result['monte_carlo_p']:.6f}"
        )
        print(f"Standard Cramér's V: {result['cramers_v']:.6f}")
        print(
            "Bias-corrected Cramér's V: "
            f"{result['bias_corrected_cramers_v']:.6f}"
        )
        print(
            "Expected cells below 5: "
            f"{result['expected_cells_below_5']}/"
            f"{result['total_cells']} "
            f"({result['percent_expected_cells_below_5']:.1f}%)"
        )
        print(
            "Minimum expected count: "
            f"{result['minimum_expected_count']:.3f}\n"
        )

    print(f"Outputs written to: {args.output_dir.resolve()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
