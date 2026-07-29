# AIFTax: Comprehensive AI Failure Taxonomy

AIFTax is a curated, source-grounded dataset of documented failures in AI-integrated software and systems. It supports research in AI reliability engineering, fault tolerance, empirical software engineering, and safety assurance.

## Overview

The dataset contains **100 AI-related failure incidents** collected from public incident databases and supporting reports. Each incident is coded across **34 columns**:

- 2 administrative and linkage fields
- 32 substantive taxonomy fields
- 10 analytical groups covering classification, propagation, detection, impact, risk, recovery, causation, and timing

The four principal analytical variables are:

- **Failure Category**
- **Propagation Reach**
- **Risk Assessment**
- **Recovery Complexity**

## Replication-Package Structure

```text
Final Corpus_and Adjudicated_Log/
    AIFTax_Final_Adjudicated_Corpus.xlsx
    A

Annotation Materials/
    AIFTax_Coding_Manual.docx
    AIFTax_Annotator_A_Annotation_Completed.xlsx
    AIFTax_Annotator_B_Annotation_Completed.xlsx

Inter_Rater_Analysis/
    irr_results/
    aiftax_inter_rater_agreement.py
    AIFTax_IRR_README.md
Provenance/
    AIFTax_Source_Registry.xlsx

README.md
```

## Dataset Schema

The 34 columns are organized into the following groups:

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

## Data Collection and Annotation

### Source Identification

Incidents were gathered from:
*Primary Source*
- public AI incident databases;
- MIT AI Risk Repository
*Supportive Source*
- verified news and investigative reports;
- organizational and industrial disclosures;
- technical reports;
- regulatory or governmental materials; and
- peer-reviewed publications where applicable.

### Independent Annotation

Two annotators independently coded all 100 incidents using the same coding manual and controlled source set.

Only these fields were prefilled:

- `Case ID`
- `Source Link(s) / Citation(s)`

All substantive fields were independently coded.

## Reproduce the Inter-Rater Analysis

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
