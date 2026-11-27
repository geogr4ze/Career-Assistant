# Dedicated Employer Monitor email delivery

Standalone Apps Script project. Do **not** copy these files into `src/` or
push them into the live SMS/call project. This isolates Gmail authorization
and the existing Twilio trigger. No credentials or real addresses belong in Git.

## Setup

1. Open the temporary Apps Script project used to verify the Gmail sending alias.
2. Replace its test code with `EmailMonitor.gs`. Keep it separate from Career Assistant's
   bound script. Set the project's timezone to America/Vancouver.
3. In Project Settings > Script Properties add:
   - EMAIL_MONITOR_FROM: the verified Gmail sending alias
   - EMAIL_MONITOR_TO: your primary Gmail address (the executing account)
   - EMAIL_MONITOR_SHEET_ID: the existing tracker ID (the part between /d/ and /edit)
4. Run `enableEmployerMonitorEmail` manually, review Google permissions and authorize.
   This creates the Email Notifications tab, marks existing events BASELINE, and installs
   a five-minute trigger. It sends nothing on initialization.
5. Check the trigger and ledger. Wait for a genuine new tracker event; verify the sender,
   email contents and ledger status before reducing other alert channels. Do not insert
   false application events in production.

The monitor sends only to the configured executing account, from a verified alias.
It reads verified application submissions and recognized employer responses; it does
not contact employers, change applications, book interviews or send approvals.
Priority employer matching is an explicit illustrative list, not a definitive classification.
It links to the tracker rather than embedding private documents or raw employer email.
It sends up to ten messages per run and thirty attempts per Vancouver calendar day,
subject to Google quota. Email can arrive during existing SMS/call quiet hours.
A SENT state means Gmail accepted the request, not confirmed inbox delivery.

SENDING/REVIEW records require manual investigation and are never automatically retried.
Do not delete ledger rows to force a resend. A complete outage cannot report itself.
ChatGPT remains the approval and incident reporting channel, including system failures
not represented as application records; the new email path does not replace it yet.

Disable with `disableEmployerMonitorEmail`. Re-enabling baselines previously unseen
events, including those created during the disabled period. Existing ledger rows are
preserved. Never run multiple independent copies against the same ledger.

## Verification

Pure event extraction and rendering checks were exercised in an isolated JavaScript
environment without provider calls. Apps Script authorization, trigger execution,
ledger behavior and genuine event delivery still require a live user-side test.

Run the reproducible local checks with `node email-monitor/test.cjs`.
