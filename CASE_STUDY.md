# Case Study: OpsTrack — Recurring Task Automation

## Business problem
Operational work repeats on different cadences. Employees need one operating view, while managers need reliable pending visibility and a consistent escalation path.

## Proposed design
Google Sheets is the operating interface. Apps Script refreshes the employee console, validates proof, writes one completion event per task cycle, advances the next due date, checks overdue tasks, emails the manager, records acknowledgement, and escalates unresolved issues to leadership.

## Workflow
1. Task is defined in `Task_Master` with an owner, recurrence interval, next due date and proof type.
2. Employee opens `Employee_Console`.
3. Employee enters ticket ID for portal work or a WhatsApp update permalink/short evidence for the daily update.
4. Employee ticks `Mark Complete` once.
5. Script validates proof, records the completion and timestamp, advances cycle number and due date.
6. SLA check finds overdue tasks and sends one manager email per task cycle.
7. Manager acknowledges in the same console and adds a note.
8. If the task remains unresolved after the SLA, leadership receives one escalation email.
9. Completion closes active alert records for that cycle.

## SLA proposal
- Manager alert: first hourly SLA scan after due time (trigger timing is approximate).
- Manager/employee resolution window: 24 hours from the first manager alert.
- Leadership escalation: after 24 hours if unresolved, whether or not the manager acknowledged.
- Deduplication: alert key = task ID + cycle number; one alert per alert type per cycle.
- Operational note: Apps Script time triggers are approximate even when configured hourly; confirm that the expected timing meets business requirements.

## Daily task proof
Use a WhatsApp message permalink when available; otherwise record a brief update summary plus timestamp in the proof field. The system records employee-submitted proof but does not independently verify WhatsApp content. A future production integration could use an approved API or manager spot checks.

## Controls
- Completion requires proof.
- Duplicate completion event IDs are rejected.
- Manager and leadership alerts are deduplicated by cycle.
- Configurable recipients and SLA.
- Test with dummy tasks and owned test inboxes before enabling real recipients.

## Evaluation checklist
- [ ] 7-day task advances by seven days after completion.
- [ ] Employee updates proof in one console.
- [ ] Daily update accepts non-ticket proof.
- [ ] Overdue task sends manager email.
- [ ] Manager acknowledgement timestamp is recorded.
- [ ] Unresolved task escalates to leadership after SLA.
- [ ] Re-running checks does not duplicate alerts.
- [ ] Repeated completion does not create a second completion record.

## Limitations
This is a working prototype design requiring deployment in the user's Google account. Email triggers and user identity behavior depend on Google account authorization and Workspace configuration. It is not connected to company systems or WhatsApp.
