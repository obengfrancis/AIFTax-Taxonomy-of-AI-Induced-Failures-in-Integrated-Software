## AIFTAX operational schema: nine analytical groups and 32 substantive fields

| Group | Fields | Purpose |
|---|---|---|
| **Administrative / Linkage** (2 fields) | Case ID; Source Link(s) / Citation(s) | Links records to the controlled evidence supplied for annotation; these fields are not analytical dimensions. |
| **Core Classification** (6 fields) | Date; Entity; System Type; Title; Industry; Application Domain | Describes the incident, affected system, organization, and deployment context. |
| **Failure Classification** (4 fields) | Failure Category; Failure Mode(s); Manifestation Pattern; Manifestation—Brief Description | Separates causal origin and mechanism from the observable expression of the failure. |
| **Propagation Analysis** (5 fields) | Propagation Reach; Boundary Transfer; Transfer Mode; Amplification; Propagation Chain | Characterizes how the failure crossed boundaries, how far it reached, and whether its effects expanded. |
| **Detection Analysis** (3 fields) | Detection Timing; Detector; Detection Signal | Identifies when, by whom or what, and through which evidence the failure became observable. |
| **Impact Assessment** (1 field) | Impact Dimension(s) | Records the direct documented consequences of the failure. |
| **Risk Assessment** (3 fields) | Risk Assessment; Risk Domain; Risk Subdomain | Captures the magnitude and primary domain of documented harm. |
| **Recovery and Maintenance** (3 fields) | Recovery Complexity; Recovery Evidence Status; Recovery / Maintenance Action(s) | Captures remediation effort, evidentiary support, and required or documented maintenance actions. |
| **Causal Analysis** (4 fields) | Causal Location(s); Causal Mechanism; Missing Safeguard; Root Cause Description | Identifies where the causal condition resided, how it produced the failure, and which safeguard was absent or ineffective. |
| **Timing Analysis** (3 fields) | Lifecycle Phase; Temporal Pattern; Timing / Period | Describes when the causal condition arose and how the failure unfolded over time. |

*The annotation instrument contains two administrative and linkage fields plus
32 substantive fields. The two administrative fields are not analytical
dimensions and are excluded from all agreement and coverage figures reported in
the paper, which are computed over 32 fields (100 × 32 = 3,200 cells). The
controlled vocabularies for each field are provided in the `Lists` sheet of the
adjudicated dataset workbook.*