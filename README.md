# Career Assistant

A personal job-search automation and notification system built with Google Apps Script, Google Sheets, ChatGPT and Twilio.

## Why I built it

As a computer engineering student searching for co-op opportunities, I wanted a better way to manage the application process. Keeping track of job postings, tailored résumés, employer replies, interviews and deadlines across multiple platforms can become repetitive and time-consuming. I wanted to spend less time checking my inbox and spreadsheets and more time working on my projects, preparing for interviews and developing my technical skills.

I also thought it would be fun to have my own automated executive assistant: a system that could text me when an application progresses or call me when an interview invitation arrives or a decision needs my attention. Building it was an opportunity to apply software automation, API integration and defensive programming to a problem I actually encounter.

## How it works

```text
Employer communications (Gmail)     Job-search research
                 |                        |
                 +---- ChatGPT workflow --+
                              |
                              v
                 Google Sheets application tracker
                              |
                    Time-driven Apps Script
                              |
                  Validate, route and deduplicate
                              |
                         Twilio REST API
                         /             \
                       SMS       Text-to-speech calls
```

Google Sheets is the central application tracker and a lightweight database. The external ChatGPT workflow searches for opportunities, prepares tailored application materials, reviews employer responses and records verified events. Google Apps Script processes those records and handles the notification layer. Twilio delivers SMS messages and automated voice calls to the configured recipient.

The spreadsheet tracks applications, employer responses, interview options, confirmed appointments, approvals and application statistics. The notification script is deliberately separate from decisions about what to tell employers or when to accept an interview.

## Features

- Configurable SMS and voice notifications for submissions, replies, assessments, interviews, offers, rejections and approvals.
- Time-driven polling without a continuously running local computer.
- Quiet hours, daily attempt limits, configurable notification routes and an emergency disable switch.
- Stable event identifiers and a persistent notification ledger to reduce duplicate alerts.
- Separate dry-run, trial and production behavior; live delivery requires explicit arming.
- Provider acceptance/status tracking and conservative handling of ambiguous failures to avoid accidental duplicate requests.
- Generic notification content; detailed employer communications and personal data stay in the private tracker.

## Design principles

I wanted the system to automate repetitive tasks while retaining control over consequential decisions. Unknown information must be clarified rather than assumed; personal qualifications and achievements must be factual; interview times and employer correspondence require my authorization. The notification layer does not independently accept offers, submit applications or send employer emails. Rejection follow-up drafts are handled by the external workflow, not this script.

## Technologies

| Technology | Role |
| --- | --- |
| Google Apps Script / JavaScript | Scheduled processing, event validation, API requests and logging |
| Google Sheets | Application records, configuration, approval queue and notification ledger |
| Twilio Programmable Messaging and Voice | SMS and automated text-to-speech calls |
| Gmail and ChatGPT | External employer-response monitoring and application preparation |
| Node.js and clasp | Local checks and Apps Script source management |

## Project structure

```text
src/
  appsscript.json   Apps Script manifest
  Config.gs         Configuration, route definitions and validation
  Events.gs         Event extraction from application records
  Notifications.gs  Polling, deduplication, delivery and status refresh
  Setup.gs          Spreadsheet menu, initialization and triggers
  Twilio.gs         Trial and production REST API payloads
  Tests.gs          Non-delivery interview simulation
tests/
  run.cjs           Local syntax and logic checks
```

## Setup

1. Create your own private Google Sheet with `Applications`, `Responses` and `Interviews` tabs using the column layout in [Tracker setup](docs/tracker-setup.md).
2. In Sheets, open **Extensions > Apps Script**. Copy the files in `src/` into the bound project, or install [clasp](https://github.com/google/clasp), enable the Apps Script API, copy `.clasp.example.json` to `.clasp.json`, enter your own Script ID and use `clasp status` followed by `clasp push`. Do not force-push over an unrelated script. The local binding is ignored by Git.
3. Configure private Script Properties as described below. Use your own verified recipient and an SMS/voice-capable Twilio sender.
4. Run `setupCareerAssistant`, reload the sheet, and configure quiet hours, limits and notification routes in Settings. Setup starts new configurations in `DRY_RUN` with routes off. Keep the live switch off until testing is complete.
5. Use **Career Assistant > Preview notifications (no sends)** and optionally run `simulateInterviewAlert` in the script editor. Manually approve a test SMS and test call through the spreadsheet menu; these may incur charges.
6. Once satisfied, explicitly arm live notifications and install the polling trigger via the menu. Arming skips existing events to prevent a flood of historical alerts. To stop, use **Disable live notifications** and **Remove polling trigger**.

### Private Script Properties

| Property | Description |
| --- | --- |
| `TWILIO_ACCOUNT_SID` | Your Twilio account identifier |
| `TWILIO_AUTH_TOKEN` | Your private authentication token |
| `TWILIO_FROM_NUMBER` | Your Twilio sender number in E.164 format |
| `ALERT_TO_NUMBER` | Your own verified recipient number in E.164 format |
| `TWILIO_MODE` | `TRIAL` for predefined templates or `PRODUCTION` for custom content |

Setup stores the bound spreadsheet identifier in `CAREER_SHEET_ID`. The live switch is managed by the menu. Never commit actual values, share the private spreadsheet, or grant untrusted users script-editor access. Trial availability and restrictions depend on the Twilio account; production use may incur charges.

## Testing and limitations

```sh
node tests/run.cjs
```

The local checks run without contacting external APIs. Manual SMS and voice tests have worked in the original deployment; delivery through the complete scheduled employer-event workflow still needs verification. A successful API request is not itself proof of delivery. The external ChatGPT workflow and private spreadsheet are not distributed with this repository. Direct Gmail classification, calendar synchronization, interactive approvals over the phone and employer-portal submission are not implemented in this Apps Script layer.

The code is intended for personal, consented notifications—not unsolicited messaging or calling.

## References

- [Google Apps Script](https://developers.google.com/apps-script)
- [Installable triggers](https://developers.google.com/apps-script/guides/triggers/installable)
- [Twilio Messaging API](https://www.twilio.com/docs/messaging/api/message-resource)
- [Twilio Voice API](https://www.twilio.com/docs/voice/api/call-resource)
