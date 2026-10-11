# ⚡ Enterprise Operations & Invoicing Automation Suite

An automated operations hub, KPI command center, and dynamic invoice generation engine built entirely inside Google Workspace using **Google Apps Script (JavaScript)**.

---

## 🎯 Architecture Overview

This project replaces manual data entry and disjointed invoicing by connecting spreadsheet records directly to automated document synthesis and transactional email dispatch.
Google Sheet (Orders_Tasks)
│
├─► Event Trigger (onEdit Status Sync & Timestamping)
│
├─► Custom UI Menu (1-Click Run via Apps Script)
│     │
│     ├─► Template Data Injection (Dynamic cell layout)
│     ├─► PDF Engine (DriveApp blob conversion)
│     └─► Gmail API Dispatch (Auto-attaches PDF & notifies client)
│
└─► Executive Command Dashboard (Live KPI Telemetry: SUMIF / COUNTIF)


---

## 🚀 Key Modules & Capabilities

- **Executive Telemetry Dashboard:** Real-time KPI summaries for billed revenue, collection rates, and pending queues using zero-overhead formulas.
- **Dynamic Print-Ready PDF Generation:** Automatically maps line items and client records into an A4 print layout and compiles high-resolution PDFs directly into Google Drive.
- **Gmail API Integration:** Dispatches personalized client notifications with auto-generated attachments and audit trail confirmation.
- **Smart Audit Triggers:** Captures row timestamps and updates operational status chips on edit events.

---

## 🛠️ Tech Stack & Services

- **Platform:** Google Workspace / Google Sheets API
- **Language:** JavaScript (Google Apps Script ES6)
- **APIs Used:** `SpreadsheetApp`, `DriveApp`, `GmailApp`, `UrlFetchApp`
- **Data Architecture:** Formula data modeling (`QUERY`, `SUMIF`, `COUNTIF`, `ARRAYFORMULA`)

---

## 👤 Author
**Shubham S.**  
*Google Sheets & Google Apps Script Automation Developer*
