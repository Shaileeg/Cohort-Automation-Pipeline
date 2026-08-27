# Automated Cohort & Attendance Management Pipeline

## Why This Was Built
After observing how manual tracking, human errors, and fragmented processes were causing bottlenecks and friction in cohort progression, I took the initiative to build a complete automated solution. This project completely streamlines the participant lifecycle from initial registration and pre-task gating to attendance logging, milestone progression, and automated re-invites to ensure smooth, error-free operations.

## ⏱️ Development Timeline & Scope
* **Timeline:** Designed, coded, and deployed efficiently in just about **6 hours** total, moving rapidly from database structuring to a fully integrated multi-tab ETL and calendar-sync pipeline.

## 🛠️ Tools & Tech Stack Used
* **Data Layer:** Google Sheets (Multi-tab relational architecture, dynamic dropdowns, conditional formatting)
* **Scripting Language:** Google Apps Script (JavaScript)
* **Ecosystem / Integrations:** Google Calendar API (event tracking and guest management), Gmail API / MailApp (automated email notifications)
* **AI Assistance:** Google Gemini and Claude (utilized as an AI technical collaborator for code generation, syntax optimization, and debugging)

---

## 🔒 Data Privacy & Security Practices
When building and deploying automation scripts that handle participant data, privacy and data protection were prioritized:
* **Separation of Code and Data:** No real participant names, emails, or personal information were ever exposed to AI or hardcoded into the repository. All development and testing were conducted using isolated mock data (e.g., test users and dummy emails).
* **Workspace Ecosystem Security:** The pipeline operates entirely within native Google Workspace APIs (`SpreadsheetApp`, `CalendarApp`, `GmailApp`), meaning data stays securely within the cloud environment governed by Google's native security infrastructure.
* **Code Sanitization:** Before publishing to GitHub, all hardcoded IDs, active form links, and proprietary credentials were completely scrubbed, leaving only clean, modular, and reusable code templates.

---

## 🤖 Development Process & AI Collaboration

* **My Role:** I acted as the sole Lead Architect and Developer. I identified the operational friction caused by human error, mapped out the database schemas, defined the core business rules, structured the Google Sheets multi-tab layout, and directed the exact logic flow required for each automation loop.
* **How AI Was Used:** **Google Gemini** & **CLaude** was utilized as an AI technical co-pilot and collaborator throughout development. AI assisted in writing, optimizing, and debugging the JavaScript syntax for Google Apps Script, translating my functional requirements and workflow logic into clean, modular code over a rapid 6-hour build cycle.


## 🗺️ System Architecture & Workflow Pipeline

```mermaid
graph TD
    A[Registration Form] -->|User Submits Data| B(1. Registration & Onboarding)
    B -->|Track Preference Analyzed| C{Pre-Task Complete?}
    
    C -->|NO| D[5-Day Reminder Trigger<br>Scans pending users & emails them]
    C -->|YES| E(2. Pre-Task Verification<br>Status updated to COMPLETED)
    
    E --> F(3. Calendar & Invitation Automation<br>Auto-adds to Tech Check & Sessions)
    F --> G(4. Mastery Evaluation & Intern Stage<br>Evaluates attendance & engagement)
    
    G --> H{Excellence / Application}
    H -->|Applies / Qualifies| I(5. Mentor Pathway & Transition<br>System flags profile as Mentor)
    
    I --> J(6. Mentor-Led Operations<br>Assigned to future cohorts & updated permissions)

    style A fill:#4169E1,stroke:#333,stroke-width:2px,color:#fff
    style C fill:#006400,stroke:#333,stroke-width:2px,color:#fff
    style H fill:#006400,stroke:#333,stroke-width:2px,color:#fff
