# AIFTax: Comprehensive AI Failure Taxonomy

AIFTax is a screened, source-grounded dataset of documented failures in AI-integrated software and systems. It supports research in AI reliability engineering, fault tolerance, empirical software engineering, and safety assurance.

## Overview

The dataset contains **100 AI-related failure incidents** collected from public incident repositories and supporting reports. Each incident is coded across **34 columns**:

- 2 administrative/linkage fields
- 32 substantive taxonomy fields
- 9 analytical groups covering core classification, failure classification, propagation, detection, impact, risk, recovery and maintenance, causation, and timing

The four principal analytical variables are:

- **Failure Category**
- **Propagation Reach**
- **Risk Assessment**
- **Recovery Complexity**

## LLM Annotation Mapping

- Annotator A = LLM A, ChatGPT-5.6 Sol (original 100-case round)
- Annotator B = LLM B, Claude Opus 5 (original 100-case round)
- Annotator C = LLM A, ChatGPT-5.6 Sol (45-case held-out round)
- Annotator D = LLM B, Claude Opus 5 (45-case held-out round)

## Replication-Package Structure

The replication package contains the following artifact groups:

```text
Final Corpus_and Adjudicated_Log/
    Cross_Dimensional_Analysis/
    figures/
    AIFTax_Final_Adjudicated_Dataset.xlsx
    [adjudication log and related outputs]

Annotation Materials/
    AIFTax_Coding_Manual.docx
    AIFTax_Annotator_A_Annotation_Completed.xlsx
    AIFTax_Annotator_B_Annotation_Completed.xlsx
    AIFTax_Annotator_C_Annotation_Completed.xlsx
    AIFTax_Annotator_D_Annotation_Completed.xlsx
    AIFTax_Annotator_A.xlsx
    AIFTax_Annotator_B.xlsx
    AIFTax_Annotator_C.xlsx
    AIFTax_Annotator_D.xlsx
    [LLM_Annotation_Prompt.txt]

Inter_Rater_Analysis/
    irr_results/
    aiftax_inter_rater_agreement.py
    AIFTax_IRR_README.md

Held-Out Analysis/
    [Annotator C and D held-out workbooks]
    aiftax_heldout_analysis.py
    aiftax_heldout_agreement.py
    [held-out analysis outputs]

Author Audit/
    AIFTAX_Author_Audit_Protocol_Revised_2026-09-02.md
    [30-incident author-audit workbook]
    [audit sampling and analysis scripts]
    [field-level and provenance-stratified outputs]

Provenance/
    AIFTax_Source_Registry.xlsx

scrpits that run from project directory.
AIFTAX operational schema: nine analytical groups and 32 substantive fields.md
Representational gap analysis motivating the design of AIFTAX.md
README.md
```

File and directory names may be normalized in the final repository; the README should be kept synchronized with the released structure.

## Dataset Schema

The 34 columns are organized as follows:

| Group | Fields |
|---|---|
| **Administrative / Linkage** | Case ID; Source Link(s) / Citation(s) |
| **Core Classification** | Date; Entity; System Type; Title; Industry; Application Domain |
| **Failure Classification** | Failure Category; Failure Mode(s); Manifestation Pattern; Manifestation — Brief Description |
| **Propagation Analysis** | Propagation Reach; Boundary Transfer; Transfer Mode; Amplification; Propagation Chain |
| **Detection Analysis** | Detection Timing; Detector; Detection Signal |
| **Impact Assessment** | Impact Dimension(s) |
| **Risk Assessment** | Risk Assessment; Risk Domain; Risk Subdomain |
| **Recovery and Maintenance** | Recovery Complexity; Recovery Evidence Status; Recovery / Maintenance Action(s) |
| **Causal Analysis** | Causal Location(s); Causal Mechanism; Missing Safeguard; Root Cause Description |
| **Timing Analysis** | Lifecycle Phase; Temporal Pattern; Timing / Period |

The following fields may contain multiple semicolon-separated labels:

- `Failure Mode(s)`
- `Impact Dimension(s)`
- `Recovery / Maintenance Action(s)`
- `Causal Location(s)`

## Notes on the Released Files

**Notes / Flag field.** The coding manual describes an optional Notes / Flag field (field 35, outside the 32 substantive fields). It was unused by both original-round LLM annotators and is therefore omitted from the adjudicated dataset. It remains in the released annotator workbooks where applicable.

**Timing / Period.** The coding manual designates this field as optional when no supported date range or recurrence note is available. It is populated for 30 of 100 incidents; remaining cells are blank by design rather than uncoded.

## Data Collection and Annotation

### Source Identification

Candidate incidents were drawn from public AI incident repositories, including the AI Incident Database and MIT AI Risk Repository. Eligibility and annotation were supported by traceable public evidence, including where applicable:

- peer-reviewed publications;
- regulatory or governmental materials;
- investigative and accountable journalism;
- organizational or industrial disclosures;
- litigation records; and
- technical reports.

Repository entries were used primarily for discovery and source tracing when additional source detail was required.

### Independent LLM Annotation

Two LLM annotators independently coded all 100 incidents using the same frozen coding manual, controlled annotation workbook, and controlled source set. Neither model was provided the other model's annotations.

Only the following fields were prefilled:

- `Case ID`
- `Source Link(s) / Citation(s)`

All 32 substantive fields were independently coded. The original LLM workbooks were preserved unchanged before agreement analysis and source-grounded adjudication.

The annotation prompts are included in the replication package.


## Reproduce the Original Inter-Model Agreement Analysis

Requirements:

- Python 3.10 or later
- `openpyxl`

Install the dependency:

```bash
python -m pip install openpyxl
```

Run from the replication-package root:

```bash
python Inter_Rater_Analysis/aiftax_inter_rater_agreement.py \
  --annotator-a "Annotation Materials/AIFTax_Annotator_A_Annotation_Completed.xlsx" \
  --annotator-b "Annotation Materials/AIFTax_Annotator_B_Annotation_Completed.xlsx" \
  --output-dir "Inter_Rater_Analysis/irr_results"
```

Complete field-level agreement outputs, adjudication records, held-out analyses, author-audit analyses, and provenance outputs are included in their corresponding replication-package directories.

## Potential Uses

AIFTax may support:

- failure-mode and root-cause analysis
- reliability and resilience benchmarking
- AI-aware fault-tolerance design
- propagation and amplification modeling
- educational materials
- policy and governance analysis
- automated failure-classification tools
- recovery and maintenance studies
- comparative analysis with other AI incident datasets

The corpus is a curated collection of publicly documented incidents and should not be treated as statistically representative of all AI failures.

## Contact

For questions, corrections, or suggestions, open an issue in the repository or contact:

Email: `[CONTACT EMAIL]`
