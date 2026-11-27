# Delivery reliability and deployment checks

## Status refresh

- A sent SMS remains eligible for delivery-status checks.
- The oldest last-checked unresolved rows are refreshed first, up to 15 per poll.
- Three consecutive failed lookups, a missing provider resource (HTTP 404/410),
  invalid provider IDs/timestamps, or 72 hours without a terminal status move
  the row to REVIEW. The last known provider status is retained.
- REVIEW is not proof of failed delivery. Inspect Twilio's console before any
  manual action. No status lookup creates a new message or call.
- Duplicate ledger keys are quarantined in both scripts. Other events continue;
  all physical attempts count toward limits. Inspect duplicates manually rather
  than deleting historical evidence.
- Email now defers delivery during the existing tracker quiet window.

## Tests before production update

Run from the repository root:

```powershell
node tests/run.cjs
node email-monitor/test.cjs
node tests/status-regression.cjs
```

The regression suite uses synthetic tracker rows and mocked provider services.
It exercises the automatic poll functions, not just manual send helpers. It
never creates real tracker events, sends emails, or contacts Twilio.

Deploy the main `src/` changes to the existing bound project using its existing
clasp configuration, after confirming the target script. Do not re-arm the
system: existing settings, trigger, ledger, and private properties stay intact.
For the standalone email project, replace only its existing EmailMonitor.gs
with the new version, save, and verify the existing polling trigger. Do not
run the enable function again when already enabled.

## Live validation (not performed by these tests)

Review the original trial message and call in the Twilio console. Do not erase
them or infer failed delivery from an unavailable status lookup. Check both
Apps Script projects' Triggers and Executions pages. A genuine verified future
application event can establish automatic delivery end to end. For an immediate
live test, use an isolated test tracker and explicitly authorized provider sends;
never insert fictitious application events into the production tracker.
