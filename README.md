# OpsTrack — Allo Innoware AI Operations Manager Assignment

## What this prototype demonstrates
- Recurring task master with 1-, 7-, 15- and 30-day cycles
- One employee operating sheet (`Employee_Console`)
- Ticket ID or daily-task proof captured before completion is accepted
- Completion log with cycle, owner, actor, timestamp and proof
- Next due date and cycle number generated automatically after completion
- Manager email on overdue tasks
- Manager acknowledgement capture
- Leadership escalation after a configurable SLA delay
- Alert deduplication by task ID + cycle + alert type
- Menu action to run the SLA check on demand

## Important
This package is a starter project, not a deployed public Google Sheet. Google Sheets and Apps Script must be installed under an account you control, authorized, and shared with the recruiter. Never put real employee or company data into this demo without approval.

## Install
1. Create a new Google Sheet in your Google account.
2. Name it `OpsTrack - Recurring Task Automation`.
3. Go to **Extensions → Apps Script**.
4. Replace `Code.gs` with the contents of `Code.gs` in this folder.
5. Add an HTML file named `Sidebar`; paste `Sidebar.html`.
6. Save and reload the spreadsheet.
7. Choose **OpsTrack → Initialize / Refresh system** and authorize.
8. In `Config`, replace the sample manager and leadership email addresses with test recipients you control. Keep `Test Mode` as `YES`.
9. Review `Task_Master` and `Employee_Console`. Use dummy data only.
10. Use **OpsTrack → Run SLA check now** to test overdue rows.
11. After verifying the recipient and behavior, choose **OpsTrack → Install daily automation** and authorize creation of the hourly SLA trigger.
12. Share the Google Sheet with the evaluator using the access level they requested. The assignment says editor access is required. Only do this after checking that the sheet contains no private data.

## Sheet tabs
- `Task_Master`: configured task recurrence, owner, due date, required proof and latest completion.
- `Employee_Console`: primary employee interface. Proof goes in column H; mark complete in I. Manager acknowledgement is in L; note in M.
- `Completion_Log`: one record per completed task cycle.
- `Alerts_Log`: manager/leadership alert records, acknowledgement and resolution timestamps.
- `Config`: recipients and SLA settings.

## Proposed SLA
- Hourly scheduled scan, allowing a same-day manager alert after the task's due time (Google may run triggers with some timing variation).
- Manager email sent on the first scan after a task is overdue.
- Leadership escalation after 24 hours from the first manager alert if the task is still unresolved, whether or not the manager acknowledged.
- One manager alert and one leadership escalation per task cycle; completion closes alerts for that cycle.
- Hourly time triggers are approximate and not guaranteed to run at an exact minute. Confirm that this timing meets the evaluator's SLA.

## Test plan
1. **Recurring cycle:** complete a 7-day task with a ticket ID. Confirm `Task_Master` next due date moves forward seven days and cycle increments.
2. **Single-point update:** enter proof and tick completion only in `Employee_Console`; confirm `Completion_Log` records it.
3. **Daily task:** enter a WhatsApp message permalink or concise evidence in column H and complete the daily task. Do not claim the script independently verifies WhatsApp; it records employee-supplied proof.
4. **Same-day manager alert:** set a dummy task due date/time in the past, configure your own test email as manager, run SLA check. Confirm one email and one `MANAGER` row in `Alerts_Log`.
5. **Acknowledgement:** tick the manager acknowledgement checkbox and enter a note. Confirm acknowledgement timestamp appears in `Alerts_Log`.
6. **Leadership escalation:** set `Escalation Delay Hours` to `0` temporarily, run SLA check on an overdue task, and confirm one `LEADERSHIP` row/email to a test recipient. Restore the value to 24 afterward.
7. **Timing:** demonstrate manager alert on first overdue scan and escalation after the configured delay. Note that a daily trigger may run at a varying time within the selected hour.
8. **No duplicates:** run SLA check multiple times. Confirm the same task cycle has only one manager alert and one leadership escalation. Tick completion only once and confirm the event ID guard prevents duplicate completion logging.

## Known prototype limitations / review before submission
- Google Apps Script's `Session.getActiveUser().getEmail()` may be blank depending on account and authorization context; fallback is the task assignee label. For stronger identity verification, use domain-managed accounts or a controlled Google Form/app.
- The daily task proof is employee-supplied evidence. The script does not read WhatsApp or verify the post independently.
- Google time-based triggers run approximately, not at a guaranteed exact minute.
- Test-mode labels emails but does not prevent sending; use test addresses in Config until fully validated.
- The prototype is designed for a single task cycle per task in the master. If employees may miss multiple cycles without completing, add a separate cycle-instance table to track every missed period.
- Have the evaluator's expected SLA and sharing policy confirmed before production use.
