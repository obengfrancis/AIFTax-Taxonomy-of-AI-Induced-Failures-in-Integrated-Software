# Figure 1 Failure-Mode Aggregation Mapping

## Purpose

Figure 1 of the AIFTAX paper visualizes the adjudicated `Failure Mode(s)` field using
11 display groups. These display groups are a presentation-only aggregation of the
16 controlled Failure Mode labels observed in the 100-incident adjudicated corpus;
they do **not** replace or modify the underlying AIFTAX labels.

The five incidents whose primary `Failure Category` is `Mixed/Hybrid` are shown as
a separate `Distributional + Operational` bar and are excluded from the 11
failure-mode display groups to avoid double counting.

## Mapping used for Figure 1

Counts below are for the 95 non-Mixed/Hybrid incidents used in the 11 display groups.

| Primary category | AIFTAX Failure Mode(s) label | Concise coding-manual definition | Figure 1 display group | Cases |
|---|---|---|---|---:|
| Operational | Capability Limitation | System cannot perform the intended task reliably even under expected conditions. | Capability Limitation | 20 |
| Operational | Design/Implementation Defect | Software, model implementation, configuration, interface, or logic defect. | Design & Implementation | 8 |
| Operational | Objective/Metric Misalignment | Optimization target rewards behavior inconsistent with safety, policy, or user intent. | Design & Implementation | 4 |
| Operational | Infrastructure/Dependency Failure | Network, platform, sensor, database, or external dependency causes AI-enabled service failure. | Design & Implementation | 1 |
| Operational | Safety Validation/Oversight Gap | Insufficient testing, hazard analysis, review, or human oversight before or after deployment. | Safety Validation | 8 |
| Operational | Content Generation/Filtering Failure | Harmful or inappropriate content passes because safeguards or filtering are inadequate, without a distribution mismatch or attacker. | Content Generation | 4 |
| Operational | Information Quality/Input Handling Failure | Malformed, incomplete, stale, or incorrectly processed information not primarily caused by distribution shift. | Information Quality | 3 |
| Distributional | Bias & Discrimination | Systematic performance disparity associated with underrepresentation or historical inequity. | Bias & Discrimination | 24 |
| Distributional | Data/Domain Shift | Deployment conditions differ from training or validation conditions. | Prediction & Drift | 3 |
| Distributional | Representation/Low-Resource Gap | Language, population, context, or class is insufficiently represented. | Prediction & Drift | 3 |
| Distributional | Hallucination/Confabulation | Plausible but unsupported or fabricated output where the evidence supports a mismatch between deployment demand and learned or knowledge support, without deliberate attack. | Hallucination | 3 |
| Adversarial | Disinformation/Manipulation | AI is intentionally used to fabricate, distort, or amplify deceptive information. | Disinformation & Manipulation | 5 |
| Adversarial | Nation-State/Military Exploitation | State-linked or military use amplifies attack, influence, or intrusion capability. | Disinformation & Manipulation | 2 |
| Adversarial | Identity/Impersonation/Deepfake | AI creates or enables deceptive identity representation. | Identity & Targeting | 4 |
| Adversarial | Targeting/Surveillance Abuse | AI is deliberately used to identify, target, or monitor persons or groups harmfully. | Identity & Targeting | 1 |
| Adversarial | Security Exploitation | Attacker exploits vulnerabilities, access control, model integrity, or system interfaces. | Security Exploitation | 2 |

## Display-group validation

The mapping reproduces the Figure 1 counts exactly:

| Figure 1 display group | Component AIFTAX labels | Cases |
|---|---|---:|
| Capability Limitation | Capability Limitation | 20 |
| Design & Implementation | Design/Implementation Defect (8) + Objective/Metric Misalignment (4) + Infrastructure/Dependency Failure (1) | 13 |
| Safety Validation | Safety Validation/Oversight Gap | 8 |
| Content Generation | Content Generation/Filtering Failure | 4 |
| Information Quality | Information Quality/Input Handling Failure | 3 |
| Bias & Discrimination | Bias & Discrimination | 24 |
| Prediction & Drift | Data/Domain Shift (3) + Representation/Low-Resource Gap (3) | 6 |
| Hallucination | Hallucination/Confabulation | 3 |
| Disinformation & Manipulation | Disinformation/Manipulation (5) + Nation-State/Military Exploitation (2) | 7 |
| Identity & Targeting | Identity/Impersonation/Deepfake (4) + Targeting/Surveillance Abuse (1) | 5 |
| Security Exploitation | Security Exploitation | 2 |
| **Subtotal: 11 failure-mode groups** |  | **95** |
| **Mixed/Hybrid shown separately** |  | **5** |
| **Total** |  | **100** |

## Mixed/Hybrid handling

The five `Mixed/Hybrid` incidents contain one Distributional and one Operational
co-primary failure mode and are displayed separately rather than being allocated
to the 11 groups above:

| Case ID | Adjudicated Failure Mode(s) |
|---:|---|
| 35 | Bias & Discrimination; Content Generation/Filtering Failure |
| 59 | Data/Domain Shift; Safety Validation/Oversight Gap |
| 81 | Representation/Low-Resource Gap; Information Quality/Input Handling Failure |
| 90 | Bias & Discrimination; Information Quality/Input Handling Failure |
| 100 | Hallucination/Confabulation; Information Quality/Input Handling Failure |

Thus, Figure 1 contains 95 incidents in the 11 aggregated failure-mode groups plus
5 Mixed/Hybrid incidents in the separate `Distributional + Operational` bar.

## Reproducibility note

`figures.py` plots a precomputed `subcategory` value from `clean_data.json`; it
does not itself encode the 16-to-11 mapping. This file makes that aggregation
explicit so the Figure 1 transformation can be inspected without reverse
engineering preprocessing code.

The aggregation is for visualization only. Analyses using `Failure Mode(s)` should
continue to use the original adjudicated AIFTAX controlled labels.
