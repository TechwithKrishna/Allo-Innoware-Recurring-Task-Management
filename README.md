# Allo Innoware — Recurring Task Tracking Prototype

## Deliverables
- `Allo_Innoware_Recurring_Task_Prototype.xlsx`: workbook template with Task Master, Employee Work Queue, Cycle Log, Manager Dashboard, Alert Log, Config, Test Cases and setup instructions.
- `Code.gs`: Google Apps Script automation.

## Activate in Google Sheets
1. Upload the `.xlsx` file to Google Drive and open it with Google Sheets.
2. In `Config`, replace `manager@company.com` and `leadership@company.com` with approved test addresses.
3. Open **Extensions → Apps Script** and paste `Code.gs`.
4. Save, return to the sheet, reload, then choose **Allo Ops → Setup / Repair Automation**.
5. Approve requested permissions. This installs hourly triggers for alert processing and view refresh.
6. Confirm the spreadsheet timezone is `Asia/Kolkata` (File → Settings).

## Operating flow
- Employee filters `Employee Work Queue` by their name.
- For Portal Task A/B/C, enter the external Ticket ID in the proof column and tick `Complete?`.
- For Daily Operations Update, share the WhatsApp update, enter a short reference such as `Ops WhatsApp group update — 09:42`, then tick `Complete?`.
- Script writes completion actor/time/proof into `Cycle Log` and creates the next cycle. Employees do not edit the Cycle Log.
- Manager reviews `Manager Dashboard` and ticks `Manager Acknowledged?`. The script records manager identity/time.
- Hourly automation sends the manager email after due time + 15 minutes. Leadership escalation is sent once if unresolved 24 hours after due time or if manager acknowledgement is not recorded within 4 business hours after the manager alert, whichever comes first.
- `Alert Log` uses a unique key per cycle and alert type, preventing repeat sends on subsequent runs.

## SLA rationale
- 15-minute grace avoids alerts for minor timing/transmission delays.
- 4 business hours gives the manager a reasonable window to review an alert during the working day.
- 24 hours after due time is the hard maximum for unresolved work before leadership visibility.
- Alerts are sent once per cycle per escalation level. A failed email is logged as failed; correct the address/permissions and clear that failed alert row only if you want to retry. A manager acknowledgement after the 4-business-hour SLA does not erase the missed-SLA escalation condition.

## Important implementation notes
- Daily completion is employee attestation, not direct WhatsApp API verification. For stronger evidence, require a WhatsApp message permalink/screenshot reference or integrate an approved WhatsApp Business API provider.
- Recurrence is completion-anchored: a 7/15/30-day cycle is due interval days after completion. The first daily cycle is due today (or next weekday if setup happens on a weekend); subsequent daily cycles recur on the next weekday (Mon–Fri).
- Employee-entered proof is preserved by hourly view refreshes until the employee checks Complete. The queue is rebuilt by cycle ID without discarding unsent proof.
- Apps Script hourly triggers are not exact-to-the-minute schedulers. The manager alert will be processed on the first hourly trigger after the 15-minute grace. For a strict minute-level SLA, install a 5-minute trigger.
- Google Sheets `onEdit` simple trigger handles employee and manager checkboxes. The first setup authorization is required for MailApp.
- Test with dummy addresses and dummy due dates before enabling real notifications.
