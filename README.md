# Automated Cohort & Attendance Management Pipeline

## Why This Was Built
After observing how manual tracking, human errors, and fragmented processes were causing bottlenecks and friction in cohort progression, I took the initiative to build a complete automated solution. This project completely streamlines the participant lifecycle from initial registration and pre-task gating to attendance logging, milestone progression, and automated re-invites to ensure smooth, error-free operations.

## ⏱️ Development Timeline & Scope
* **Timeline:** Designed, coded, and deployed efficiently in just about **6 hours** total, moving rapidly from database structuring to a fully integrated multi-tab ETL and calendar-sync pipeline.

## 🛠️ Tools & Tech Stack Used
* **Data Layer:** Google Sheets (Multi-tab relational architecture, dynamic dropdowns, conditional formatting)
* **Scripting Language:** Google Apps Script (JavaScript)
* **Ecosystem / Integrations:** Google Calendar API (event tracking and guest management), Gmail API / MailApp (automated email notifications)
* **AI Assistance:** Google Gemini (utilized as an AI technical collaborator for code generation, syntax optimization, and debugging)

---

## What I Did (Step-by-Step Implementation)
1. **Designed the Database Architecture:** Structured a multi-tab Google Sheet to cleanly separate registration data, pre-task tracking, master intern details, monthly attendance logs, and make-up queues.
2. **Built Automated Gating & Ingestion:** Created scripts to automatically ingest form responses and lock entry into master tables until pre-tasks were successfully completed.
3. **Integrated Calendar & Track Syncing:** Implemented date and tracking logic (`WD` vs. `WE`) to automatically add participants as guests to shared Google Calendar sessions.
4. **Automated Attendance, Milestones & Mentor Conversion:** Set up daily time-based triggers to parse calendar history, build dynamic monthly attendance tabs, track a 4-class milestone to automatically trigger mentor upgrade emails, and dynamically parse Gmail inbox replies to assign roles (`Intern` vs. `Mentor`) and track preferences (`WD` vs. `WE`).
5. **Engineered Exception & Make-Up Loops:** Built automated parsing for inbox replies and absence tracking that dynamically routes absentees into a make-up queue, updates their roles, and re-invites them to upcoming sessions.

---

## 🔄 End-to-End Workflow Architecture

1. **Registration & Pre-Task Gating:** 
   - Participants register and submit pre-tasks. 
   - Scripts (`RegistrationAutomation.gs`, `PreTaskAutomation.gs`, `PreTaskTracker.gs`) filter submissions, ensuring users only enter the master tracking sheet (`Intern Details`) once their pre-tasks are verified.
2. **Calendar Scheduling & Track Preferences:** 
   - Date logic (`DateCalculator.gs`) and calendar automation (`CalendarInvites.gs`) read participant tracks (Weekday `WD` or Weekend `WE`) and auto-sync them as guests to shared Google Calendar session links.
3. **Attendance & Milestone Tracking:** 
   - Daily triggers (`AttendanceTracker.gs`) parse calendar guest lists from the past 24 hours and dynamically log attendance into monthly tabs (`PRESENT` or `ABSENT`).
   - `MentorManagement.gs` monitors intern completion milestones. Once an intern hits **4 completed classes**, it auto-updates their status to `Invited` and sends an invitation email.
4. **Mentor Response Parsing & Unified Guest Sync:** 
   - `MentorResponses.gs` scans the inbox for mentor invitation replies, parses acceptance/decline status and track choices (`WD`/`WE`), updates sheet columns, and automatically appends both old and new participants as guests (`event.addGuest(email)`) to shared master calendar events so everyone shares the exact same Google Meet link.
5. **Absence & Make-Up Workflow:** 
   - `ReinviteAutomation.gs` captures anyone marked `ABSENT`, routes them into a dedicated `Reinvite` tracking tab with their specific role, sends a make-up notification email, and re-syncs them to upcoming calendar sessions.

---

## 📁 Script Modules Breakdown

| File Name | Description |
| :--- | :--- |
| `DateCalculator.gs` | Manages date logic and timeline calculations for sessions. |
| `RegistrationAutomation.gs` | Ingests and processes raw registration form data. |
| `PreTaskAutomation.gs` & `PreTaskTracker.gs` | Validates pre-task completion and gates entry into master tables. |
| `CalendarInvites.gs` | Automatically syncs participants to weekday/weekend calendar events. |
| `AttendanceTracker.gs` | Runs nightly triggers to build monthly attendance tabs and log presence. |
| `MentorManagement.gs` | Monitors the 4-class milestone to trigger mentor upgrade invitations. |
| `MentorResponses.gs` | Parses Gmail inbox replies for mentor acceptances/preferences and syncs all roles into unified shared calendar session links. |
| `ReinviteAutomation.gs` | Captures absentees, populates the make-up tab, and re-invites users. |

---

## 🔒 Data Privacy & Security Practices
When building and deploying automation scripts that handle participant data, privacy and data protection were prioritized:
* **Separation of Code and Data:** No real participant names, emails, or personal information were ever exposed to AI or hardcoded into the repository. All development and testing were conducted using isolated mock data (e.g., test users and dummy emails).
* **Workspace Ecosystem Security:** The pipeline operates entirely within native Google Workspace APIs (`SpreadsheetApp`, `CalendarApp`, `GmailApp`), meaning data stays securely within the cloud environment governed by Google's native security infrastructure.
* **Code Sanitization:** Before publishing to GitHub, all hardcoded IDs, active form links, and proprietary credentials were completely scrubbed, leaving only clean, modular, and reusable code templates.

---

## 🤖 Development Process & AI Collaboration

* **My Role:** I acted as the sole Lead Architect and Developer. I identified the operational friction caused by human error, mapped out the database schemas, defined the core business rules, structured the Google Sheets multi-tab layout, and directed the exact logic flow required for each automation loop.
* **How AI Was Used:** **Google Gemini** was utilized as an AI technical co-pilot and collaborator throughout development. AI assisted in writing, optimizing, and debugging the JavaScript syntax for Google Apps Script, translating my functional requirements and workflow logic into clean, modular code over a rapid 6-hour build cycle.


The diagram below illustrates the end-to-end data flow and automated lifecycle engineered for the Course:

[ Registration Form ] 
         │
         ▼ (User Submits Data)
 ┌──────────────────────────────────────────────┐
 | 1. REGISTRATION & ONBOARDING                 |
 |    • Google Sheets logs submission           |
 |    • Track preference analyzed (Weekday/End) |
 └──────────────────────┬───────────────────────┘
                        │
                        ▼ (Pre-Task Complete?)
        ┌───────────────┴───────────────┐
        │ NO                            │ YES
        ▼                               ▼
 ┌───────────────┐               ┌──────────────────────────────────────┐
 | 5-DAY REMINDER|               | 2. PRE-TASK VERIFICATION             |
 | Script scans  |               | • Checkbox marked TRUE               |
 | pending users |               | • Status updated to "COMPLETED"      |
 | & emails them |               └──────────────────┬───────────────────┘
 └───────────────┘                                  │
                                                    ▼
                                 ┌──────────────────────────────────────┐
                                 | 3. CALENDAR & INVITATION AUTOMATION  |
                                 | • Auto-adds to Tech Check & Sessions |
                                 | • Sends confirmation email & links   |
                                 └──────────────────┬───────────────────┘
                                                    │
                                                    ▼ (Course Completion)
                                 ┌──────────────────────────────────────┐
                                 | 4. MASTERY EVALUATION & INTERN STAGE |
                                 | • Evaluates attendance & engagement  |
                                 | • Confirms course completion         |
                                 └──────────────────┬───────────────────┘
                                                    │
                                                    ▼ (Excellence / Application)
                                 ┌──────────────────────────────────────┐
                                 | 5. MENTOR PATHWAY & TRANSITION       |
                                 | • High-performing interns apply/opt  |
                                 |   in for the Mentor Track            |
                                 | • System flags them as "Mentor"      |
                                 |   in the database                    |
                                 └──────────────────┬───────────────────┘
                                                    │
                                                    ▼
                                 ┌──────────────────────────────────────┐
                                 | 6. MENTOR-LED OPERATIONS             |
                                 | • Assigned to track future cohorts   |
                                 | • Script updates calendar invites to |
                                 |   include them as co-hosts/mentors   |
                                 | • Feedback loops track mentor impact |
                                 └──────────────────────────────────────┘