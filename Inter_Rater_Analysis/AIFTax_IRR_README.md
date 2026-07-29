# AIFTax Inter-Rater Agreement Script

## Purpose

`aiftax_inter_rater_agreement.py` compares two independent AIFTax annotation workbooks by `Case ID`, verifies alignment, and calculates agreement before adjudication.

The principal analytical variables are:

1. `Failure Category` — nominal Cohen's kappa
2. `Propagation Reach` — quadratic weighted Cohen's kappa
3. `Risk Assessment` — quadratic weighted Cohen's kappa
4. `Recovery Complexity` — quadratic weighted Cohen's kappa

`Case ID`, `Source Link(s) / Citation(s)`, and `Notes / Flag (optional)` are also excluded from the agreement calculation.

## Requirements

- Python 3.10 or newer
- `openpyxl`

Install the dependency with:

```bash
python -m pip install openpyxl
```

## Run

```bash
python Inter_Rater_Analysis/aiftax_inter_rater_agreement.py \
      --annotator-a "Annotation Materials/AIFTax_Annotator_A_Annotation_Completed.xlsx" \
      --annotator-b "Annotation Materials/AIFTax_Annotator_B_Annotation_Completed.xlsx" \
      --output-dir "Inter_Rater_Analysis/irr_results"
```

The script automatically detects the annotation header row and normalizes known header spacing and spelling issues.

## Main outputs

- `AIFTax_IRR_Report.xlsx` —Workbook containing summaries, confidence intervals, disagreement lists, label distributions, validation results, and principal-variable confusion matrices
- `AIFTax_IRR_Report.json` — machine-readable complete report
- `AIFTax_IRR_Principal_Summary.csv`
- `AIFTax_IRR_Supporting_Summary.csv`
- `AIFTax_IRR_Multilabel_Summary.csv`
- `AIFTax_IRR_Principal_Disagreements.csv`
- `AIFTax_IRR_Controlled_Disagreements.csv`

## Metric treatment

- Nominal controlled fields: exact agreement and unweighted Cohen's kappa
- Ordered fields: exact agreement, unweighted kappa, linear weighted kappa, and quadratic weighted kappa
- Multi-label fields: exact set agreement, mean Jaccard similarity, micro-F1, and macro per-label kappa
- Structured narrative fields: normalized exact agreement as a diagnostic only, not as a categorical reliability statistic

For ordered fields, weighted kappa is calculated only for pairs falling within the defined ordered scale. `IE`, `NA`, `Other`, and values outside the ordered scale remain part of exact agreement and unweighted kappa but are excluded from weighted kappa.

The default 95% confidence intervals use 2,000 case-level bootstrap resamples with seed `2026`. Change them with:

```bash
--bootstrap 5000 --seed 2026
```