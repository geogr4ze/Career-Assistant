# Tracker setup

Create a private Google Sheet. Use these headers in row 1, in the order listed. The script checks specific header positions and expects submission dates to be real spreadsheet dates, not text.

**Applications** (A:X)

```text
Application ID | Employer | Position | Field | Location | Work mode | Duration | Posting URL | Deadline | Date found | Status | Date submitted | Channel | Confirmation / proof | First employer response | Response type | Interview requested date | Offer date | Follow-up due | Next action | Résumé link | Cover letter link | Last updated | Notes
```

**Responses** (A:K)

```text
Response ID | Application ID | Received date/time (PT) | Sender | Response type | Subject | Summary | Gmail thread / source URL | Action required | Action deadline (PT) | Action status
```

**Interviews** (A:P)

```text
Interview ID | Application ID | Employer | Position | Round / type | Employer offered time | Employer time zone | Start option (PT) | End option (PT) | Selection status | Confirmed start (PT) | Confirmed end (PT) | Meeting link / location | Recruiter | Source / thread URL | Notes
```

Split the header lines on ` | ` into separate cells. Use stable unique IDs. Submission alerts require a valid submission date and confirmation/proof. Responses must refer to an existing Application ID. Supported response types are `Acknowledgment only`, `Recruiter reply`, `Assessment`, `Interview invitation`, `Offer`, `Rejection`, `Other substantive reply`, and `Progression`. Only verified responses should be added. The initialization function creates `Settings`, `Notifications`, `Approval Queue`, and `System Logs`. Add one approval row per genuine unresolved decision with status `Pending`.

The original deployment uses Vancouver time; the code and manifest use `America/Vancouver`. Set the sheet timezone to match. A dashboard can be created separately using actual submission and milestone dates. The script does not create the original private dashboard or populate historical applications.
