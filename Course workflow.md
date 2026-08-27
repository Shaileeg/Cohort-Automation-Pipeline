# Partnership Course Automation — Complete Workflow Reference

## 1. Registration & Prerequisite Check
**File:** `RegistrationAutomation.gs`

Someone submits the Registration Form. The script checks three prerequisites (Design Thinking, Networking, Networking Live).
- **Pass** → status `QUALIFIED`, pre-task email sent with a deadline calculated by `DateCalculator.gs`.
- **Fail** → status `PREREQ_FAILED`, polite rejection email. No automated re-apply path.

## 2. Pre-Task Gate — One-Time Only
**Files:** `PreTaskAutomation.gs`, `PreTaskTracker.gs`

Nobody enters `Intern Details` (the master tracking sheet) until their pre-task is matched against a registration record. This is a **one-time requirement per person, forever** — even if someone later misses a meeting and gets rolled into a future cohort, they are never asked to redo it.

On match: registration marked `COMPLETED`, they're added to all 4 weekly sessions + the tech check (with a guaranteed Google Meet link — see §3), and they get a confirmation email. A daily reminder job nags anyone who hasn't submitted, 5 days before their cohort starts (uses the real submission link now, not a placeholder).

## 3. Sessions — Thursday (Weekday) / Sunday (Weekend) Only
**Files:** `PreTaskTracker.gs`, `CalendarInvites.gs`, `MentorResponses.gs`

- Weekday sessions: 11:00–13:00 UTC (4:45pm–6:45pm Nepal time)
- Weekend sessions: 14:00–16:00 UTC (7:45pm–9:45pm Nepal time)
- Attendance is a **manual checkbox grid** a mentor fills in — not pulled from Calendar RSVPs.
- Every event is created via the **Calendar Advanced Service**, so a Google Meet link is guaranteed regardless of anyone's personal account settings.
- Attendance/reinvite processing runs on a **day-aware weekly trigger** (`createScheduledAttendanceTriggers()`), not daily — it calculates which of the 4 meeting numbers today actually is and no-ops safely on any other day.

## 4. Miss ANY Meeting → Removed This Month, Rolled to Next
**File:** `AttandanceTracker.gs` (`processMonthlyAttendance`, `rollInternToNextCohort`)

Mastery requires attending all 4 meetings, so missing even one — first, second, doesn't matter — disqualifies that cohort immediately. The moment this happens:
1. Recorded in the attendance sheet (Column P) with which meeting caused it.
2. `removeFromRemainingSessionsThisMonth()` pulls them off the rest of this month's calendar events.
3. `rollInternToNextCohort()` re-adds them to next month's 4 sessions + tech check (no pre-task needed) and emails them.
4. Capped at **3 total invites** (tracked in `Intern Details` Column J/K) — after that, status is set to a clear final `"Stopped - no response after 3 invites"` instead of continuing to email indefinitely.
5. Once ineligible, they stop receiving per-meeting feedback emails, and the meeting-level make-up reinvite path (`ReinviteAutomation.gs`) skips them too, since a make-up can't restore eligibility anyway.

## 5. Mastery Evaluation (Meeting 4)
**File:** `AttandanceTracker.gs`

At Meeting 4: `MASTERED` requires all 4 attended + good tech ≥2x + active participation ≥2x. Either way, a final summary email goes out (this always fires, even for someone who was disqualified back at Meeting 1). If mastered, they're marked `"Eligible"` in `Intern Details` for a mentor invite.

## 6. Mentor Pathway
**File:** `MentorResponses.gs`

- Eligible graduates get invited; a daily job parses Gmail replies for accept/decline + track preference.
- **Accept** → added as guest to the shared calendar event (created if it doesn't exist yet, with a guaranteed Meet link), and `grantMentorTabAccess()` fires automatically — granting editor access **only to the separate attendance spreadsheet**, never the master file with everyone's contact info.
- **Decline** → re-invited automatically the following month via `refreshDeclinedMentorInvites()` (needs its own monthly trigger — `createMonthlyMentorReinviteTrigger()`), capped at **3 total invites**, then marked `"Declined - stopped after 3 invites"`.

## Data Architecture
- **Master spreadsheet**: `Intern Details`, `Registration Form`, `Pre-task Form`, `Reinvite`
- **Attendance spreadsheet** (`1OsVIi_rUivLyZfXABOZr8gb8Yc1Pyz0d5dOkPA_OJgc`): monthly attendance tabs only — this is what mentors get access to
- **Intern Details extra columns**: I = mentor invite attempts, J = course reinvite attempts, K = course reinvite status
- **Attendance sheet extra column**: P (16) = eligibility status

---

## ✅ Remaining Manual Setup (nothing left in code — these are one-time actions in the Apps Script/Calendar UI)

1. **Delete `MentorManagement.gs`** from the project if it's still there — consolidated into `MentorResponses.gs`.
2. **Enable the Calendar Advanced Service**: Editor → Services (+) → Google Calendar API → confirm it's named `Calendar`.
3. **Confirm the project's timezone** is Asia/Kathmandu (Project Settings → Time zone) — the trigger hours assume this.
4. **Run `createScheduledAttendanceTriggers()`** once — installs the 4 Thursday/Sunday triggers for attendance + reinvite checking.
5. **Delete any old daily trigger(s)** pointing at `runWeekdayMeetingN()`/`runWeekendMeetingN()` or the scheduled functions — no daily trigger should remain.
6. **Run `createMonthlyMentorReinviteTrigger()`** once — schedules the monthly mentor decline sweep for the 1st of each month.
7. **Confirm edit access** to the attendance spreadsheet (`1OsVIi_...`) for whichever account runs the script — needed for `grantMentorTabAccess()` to share it out.
8. **Do one dry run** of `runSetupCurrentMonthWeekday()`/`runSetupCurrentMonthWeekend()` before real participant data flows through, to confirm the tabs get created correctly in the new attendance sheet.

Everything else — pre-task gating, calendar creation with guaranteed Meet links, mastery evaluation, the miss-a-meeting rollover, and the mentor accept/decline loop — is fully wired in code and needs no further manual steps to function.