# QA Review — Allo Innoware Prototype

## Result
The workbook structure and Apps Script source pass the local checks described below. These are source-level and logic simulations, not a live Google Sheets deployment test.

## Checks performed

| Test | Local result | Evidence / scope |
|---|---|---|
| 1. Recurring cycle | PASS | Simulated completion-anchored 7-, 15- and 30-day due-date calculations. |
| 2. Single-point completion | PASS with deployment check pending | Employee Work Queue captures proof and completion checkbox; Apps Script writes the proof/completion metadata to the cycle record. Must test actual onEdit in Google Sheets. |
| 3. Daily task | PASS | Daily completion accepts a WhatsApp update reference; next weekday calculation checked (Friday → Monday). This is employee attestation, not WhatsApp API verification. |
| 4. Same-day manager alert | PASS with deployment check pending | Source checks due time + 15-minute grace and sends a manager email once per cycle. Actual email delivery requires configured address and authorization. |
| 5. Manager acknowledgement | PASS with deployment check pending | Dashboard checkbox writes acknowledgement, actor and timestamp to Cycle Log. Must test actual onEdit in Google Sheets. |
| 6. Leadership escalation | PASS with deployment check pending | Source escalates unresolved tasks after the earlier of the 4-business-hour acknowledgement SLA or 24-hour hard limit. Actual email delivery requires configuration. |
| 7. SLA timing | PASS | Local calculations checked for 15-minute manager grace, 4 business hours and 24-hour cap. Time trigger cadence affects actual delivery time. |
| 8. Duplicate prevention | PASS at logic level | Unique cycle + alert-level keys and open-cycle checks prevent repeated sends/cycles in normal sequential execution. Concurrent live execution still needs deployment testing. |

## Additional checks
- Apps Script JavaScript syntax: PASS (`node --check`).
- Workbook tabs: PASS (Instructions, Task Master, Employee Work Queue, Cycle Log, Manager Dashboard, Alert Log, Config, Test Cases).
- Configured recurrence intervals: PASS (1, 7, 15, 30 days).
- Proof types: PASS (daily confirmation and Ticket ID).
- Employee proof preservation during hourly refresh: implemented and source-reviewed.

## Not yet verified in a real Google account
1. Importing the workbook and converting it to native Google Sheets.
2. Apps Script authorization and time-trigger installation.
3. Real employee/manager checkbox edits and identity capture.
4. Actual manager and leadership email delivery.
5. End-to-end duplicate behaviour under simultaneous edits/trigger execution.

## Important deployment caveats
- Time-driven Apps Script triggers run on their schedule; with hourly triggers, an alert may be delayed until the next hourly execution. For tighter SLA handling, change both `.everyHours(1)` triggers to `.everyMinutes(5)` (if available to the account) and test quotas.
- `Session.getActiveUser().getEmail()` may be blank for some consumer-account or cross-domain contexts. In that case the code falls back to the assigned employee name. A Google Workspace deployment with suitable domain policy is preferable for reliable identity capture.
- Daily WhatsApp proof is a human attestation/reference. Automatic verification requires an approved WhatsApp Business integration.
- Email delivery and leadership recipients must be configured in the Config tab before enabling live alerts.
