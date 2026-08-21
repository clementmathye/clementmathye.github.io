# Market Analyst Operations & KPI Audit Hub

<img width="1343" height="475" alt="image" src="https://github.com/user-attachments/assets/9fb1ee57-fefa-40ac-b79a-21d249f66b2e" />


An enterprise-grade operational data model and executive audit dashboard built in Microsoft Excel to automate CRM quality assurance, track analyst deliverables, and manage workforce attendance compliance.

---

## 🎯 Key Objectives & Business Architecture

Designed as an end-to-end proof-of-work model addressing the 3 core operational pillars of Market Analyst Management:

1. **CRM Data Quality Control:** Eliminates manual inspection by automating missing-data detection, format validation, and account discrepancy flagging.
2. **Analyst Performance Governance:** Measures weekly deliverable volume against quotas, evaluates on-time delivery rates, and triggers automated escalation alerts.
3. **Workforce Attendance Tracking:** Monitors a 20-day rolling attendance grid to calculate operational capacity, handle leaves, and isolate chronic absenteeism.
4. **Executive Decision Support:** Consolidates multi-sheet transactional logs into an interactive SaaS-style KPI executive dashboard.

---

## 📊 Core Spreadsheet Architecture

```text
market-analyst-operations-hub/
│
├── Market_Analyst_Operations_Audit_Model.xlsx   <-- Core Multi-Sheet Excel Engine
│   ├── Sheet 1: Executive_KPI_Summary          <-- High-level KPI Cards & Slicers
│   ├── Sheet 2: CRM_Data_Quality_Audit         <-- Automated Account QA & Flags
│   ├── Sheet 3: Analyst_Deliverables_Tracker    <-- Output, SLA, & Target Tracking
│   └── Sheet 4: Workforce_Attendance_Grid      <-- 20-Day Operations Grid
│
├── Analyst_KPI_Executive_Summary.pdf           <-- 1-Page Executive PDF Report
├── LICENSE
├── .gitignore
└── README.md
🛠️ Advanced Excel Formula Modeling
The model utilizes robust dynamic formulas, data validation arrays, and automated logic layers:

Automated CRM Discrepancy Flagging:

Excel
=IF(OR(ISBLANK([@AccountID]), ISBLANK([@LeadStage]), ISBLANK([@AssignedAnalyst])), "FLAG: INCOMPLETE", IF([@AuditScore]<80, "FLAG: RE-AUDIT", "QA PASSED"))
On-Time Delivery SLA Calculation:

Excel
=IF([@ActualDeliveryDate]<=[@TargetDeadline], "On-Time", "Overdue / Escalated")
Dynamic Capacity & Attendance Rate %:

Excel
=COUNTIFS(Attendance_Grid[AnalystID], [@AnalystID], Attendance_Grid[Status], "Present") / 20
Dynamic Metric Aggregation via XLOOKUP & SUMIFS:

Excel
=SUMIFS(Deliverables[CompletedReports], Deliverables[AssignedAnalyst], [@AnalystName], Deliverables[Status], "Approved")
🚀 Tech Stack & Core Competencies
Tooling: Microsoft Excel (Advanced Formulas, Dynamic Named Ranges, Data Validation, Conditional Formatting Hierarchy)

Core Methodologies: Quality Assurance (QA) Auditing, SLA Governance, Workforce Capacity Planning, Executive Dashboard Design
