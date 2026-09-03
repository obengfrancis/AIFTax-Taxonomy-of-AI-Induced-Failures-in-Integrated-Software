## Representational gap analysis motivating the design of AIFTAX

| Aspect | AIID [1] | MIT Risk Repository [2] | Microsoft Agentic [3] | AIAG/VDA Extended [4] | AIFTAX design response |
|---|---|---|---|---|---|
| **Primary focus** | Incident documentation and contributed taxonomies | Cross-framework risk synthesis | Agentic AI safety and security | Functional safety of embedded-AI systems | SE analysis of documented failures in integrated software |
| **Failure classification** | Multiple harm and causal taxonomies | Causal and domain taxonomies | Agentic safety and security failure modes | Functional output failure modes | Four causal categories and supporting failure modes |
| **Risk severity** | Harm-oriented fields; no common five-level scale | No incident-level severity scale | Safety and security impact framing | Domain-specific FMEA assessment | Five-level incident risk scale |
| **Propagation and detection** | Partially represented across separate taxonomies | Not incident-level | Agent and trust-boundary interactions | Functional chains within system design | Boundary transfer, reach, amplification, and detection |
| **Recovery and maintenance** | Not systematically integrated | Not systematically represented | Mitigation guidance rather than incident recovery coding | Design-time prevention and control | Five-level recovery complexity and maintenance actions |
| **Integration perspective** | Broad AI incidents | Broad AI risks | Agentic component and tool interactions | Physical and embedded-AI components | AI–software boundary failures across deployment domains |

*This table records representational scope — whether each framework explicitly
represents these aspects — rather than empirical performance or overall utility.
The comparison is informed by established dependability and defect/anomaly-classification
concepts, including fault–error–failure relationships, defect causes and
manifestations, and detection and recovery [5], [6], [7].*

**Sources**

1. Partnership on AI. AI Incident Database. https://incidentdatabase.ai/
2. Slattery et al. The AI Risk Repository: A Comprehensive Meta-Review, Database, and Taxonomy of Risks from Artificial Intelligence. arXiv:2408.12622, 2024.
3. Microsoft AI Red Team. Taxonomy of Failure Modes in Agentic AI Systems, version 2.0. White paper, Microsoft, April 2026.
4. Campean, Yildirim, Korsunovs, and Doikin. Extending the Function Failure Modes Taxonomy for Intelligent Systems with Embedded AI Components. Proceedings of the Design Society, 4:1949–1958, 2024.
5. IEEE Standard Classification for Software Anomalies. IEEE Std 1044-2009.
6. Avizienis, Laprie, Randell, and Landwehr. Basic Concepts and Taxonomy of Dependable and Secure Computing. IEEE TDSC, 1(1):11–33, 2004.
7. Chillarege et al. Orthogonal Defect Classification. IEEE TSE, 18(11):943–956, 1992.