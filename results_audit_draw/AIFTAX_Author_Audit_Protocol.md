# AIFTAX Author Audit Protocol — Revised Reporting Version

**Original audit date:** 2026-08-27  
**Revision date:** 2026-09-02  
**Revision scope:** This revision preserves the sampling, coding, evidence, and analysis procedures used for the author audit, corrects terminology and minor wording errors, and aligns the interpretation of the 70% threshold with the manuscript. The 70% value is treated as a reporting threshold for identifying weaker constructs, not as an automatic trigger for revising the dataset or conducting a broader review.

## Purpose

This audit measures agreement between an author coding pass and the final adjudicated AIFTAX dataset. It provides an additional check on reproducibility, including sampled cells that entered the final dataset through LLM consensus. The audit covers all 32 substantive AIFTAX fields.

The audit is measurement-only: discrepancies observed in the sampled cases are reported but are not used to selectively alter those 30 records.

## Sampling

**Sampling seed:** `20260827`

Thirty incidents were drawn at random without replacement from the final 100-incident adjudicated dataset. The complete Case ID list was sorted in ascending order before sampling. The same 30 incidents were used for all 32 substantive fields, yielding **960 audited incident-field cells**.

The audit workbook contained the Case ID, controlled Source Link(s), the 32 substantive audit fields, and an optional audit-notes field. It did **not** contain the LLM annotations, adjudicated labels, disagreement indicators, cell-provenance indicators, or other fields that could reveal the final coding during the audit.

## Audited fields

The author independently coded all 32 substantive fields using the frozen coding manual and controlled vocabularies.

### Core Classification (6)
1. Date
2. Entity
3. System Type
4. Title
5. Industry
6. Application Domain

### Failure Classification (4)
7. Failure Category
8. Failure Mode(s)
9. Manifestation Pattern
10. Manifestation — Brief Description

### Propagation Analysis (5)
11. Propagation Reach
12. Boundary Transfer
13. Transfer Mode
14. Amplification
15. Propagation Chain

### Detection Analysis (3)
16. Detection Timing
17. Detector
18. Detection Signal

### Impact Assessment (1)
19. Impact Dimension(s)

### Risk Assessment (3)
20. Risk Assessment
21. Risk Domain
22. Risk Subdomain

### Recovery and Maintenance (3)
23. Recovery Complexity
24. Recovery Evidence Status
25. Recovery / Maintenance Action(s)

### Causal Analysis (4)
26. Causal Location(s)
27. Causal Mechanism
28. Missing Safeguard
29. Root Cause Description

### Timing Analysis (3)
30. Lifecycle Phase
31. Temporal Pattern
32. Timing / Period

IE/NA remained permitted where allowed by the coding manual. Timing / Period remained optional when the manual permitted no supported value.

## Evidence condition

The author used the frozen AIFTAX coding manual and the same controlled source set used for final adjudication. For each sampled incident, the author re-read the supplied evidence before assigning values and did not consult either LLM's annotation, the final adjudicated label, or cell provenance during coding.

## Analysis

Author-audit values were compared with the final adjudicated values only after the audit workbook had been frozen.

### Principal variables

For the four principal variables—Failure Category, Propagation Reach, Risk Assessment, and Recovery Complexity—the analysis reports:

- exact agreement;
- unweighted Cohen's kappa;
- quadratic-weighted Cohen's kappa for Propagation Reach, Risk Assessment, and Recovery Complexity; and
- 95% bootstrap confidence intervals using 2,000 nonparametric bootstrap resamples and the fixed bootstrap seed used in the corresponding agreement analysis.

### Other controlled single-label fields

For other controlled categorical fields, the analysis reports exact agreement and the corresponding nominal or ordinal Cohen's kappa. For ordinal fields, quadratic-weighted kappa may additionally be reported where applicable. Where severe marginal imbalance makes Cohen's kappa difficult to interpret, PABAK may be reported only as a supplementary statistic.

### Multi-label fields

For Failure Mode(s), Impact Dimension(s), Recovery / Maintenance Action(s), and Causal Location(s), the analysis reports per-label Cohen's kappa summarized by the macro average across labels. Exact-set agreement may be used as a descriptive supplementary statistic but does not replace the per-label agreement analysis.

### Structured-text and narrative fields

For structured-text and narrative fields, normalized exact agreement is used only as a diagnostic. Differences in source-grounded wording are not interpreted as categorical disagreement. Discrepancies may be inspected qualitatively to distinguish wording variation from substantive coding differences.

### Recovery-evidence diagnostic

For Recovery Evidence Status, the analysis reports exact agreement and unweighted Cohen's kappa. It also reports the collapsed comparison relevant to Recovery Complexity: whether each incident received an ordinal Recovery Complexity rating versus IE/NA.

### Cell-provenance analysis

Results are additionally stratified by **cell provenance**: LLM-consensus cells versus author-adjudicated cells. Provenance is determined separately for each sampled incident-field cell from the two unchanged pre-adjudication LLM workbooks.

Because some field-by-provenance strata are small, provenance-specific results are interpreted descriptively and accompanied by counts where appropriate.

## Reporting and interpretation rule

The audit is measurement-only. Discrepancies observed in the sampled cases are reported but are not used to selectively alter those 30 records.

For the four principal variables and Transfer Mode, **70% exact agreement is used only as a reporting threshold for identifying weaker constructs**. Falling below this threshold does not automatically trigger revision of the adjudicated dataset or a broader review. Instead, it indicates that claims depending materially on the affected construct should be interpreted cautiously and, where relevant, qualified in the paper.

For the remaining controlled and multi-label fields, agreement below 70% is treated as a diagnostic indication of weaker reproducibility. Such results may motivate targeted inspection or qualification of claims that materially depend on those fields, but do not imply revision of unrelated findings.

No fixed 70% threshold is applied to narrative or structured-text fields because normalized wording agreement is not a direct measure of categorical reproducibility.

Recovery Evidence Status is treated diagnostically. Disagreement on this field is used to qualify interpretation of Recovery Complexity, particularly the Inferred-versus-IE boundary, rather than to trigger revision of a headline distribution.

Any broader follow-up review is therefore motivated by the observed pattern and substantive importance of disagreement, not by an automatic threshold.
