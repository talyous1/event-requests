# HQA Spring Event & Volunteer Requests

Teachers request approval for class parties, school events, field trips, and parent or student volunteers/observers. The Principal approves, asks for changes, or denies from a dashboard. Everything is saved to a Google Sheet and emails go out automatically.

| File | What it is |
|---|---|
| `Common.gs`, `Logo.gs`, `WebApp.gs`, `Index.html`, `Dashboard.html`, `Styles.html`, `appsscript.json` | The app: teacher request page + Principal dashboard |
| `GoogleForm.gs` | Optional alternative to `WebApp.gs` + `Index.html`: builds a Google Form for teachers. **Use one or the other, never both in the same project** |
| `event-requests-trial.html` | Click-through demo with sample data. Download it and open it in a browser |

> GitHub only stores the code. The app runs in Google Apps Script. Keep this repo **private**.

## Set up

1. Make a new Google Sheet, then go to **Extensions > Apps Script**.
2. Create a file for each app file above (same names) and paste the code in. Skip `GoogleForm.gs` and the trial page.
   To edit `appsscript.json`, turn on **Project Settings > Show "appsscript.json"**.
3. Check `CONFIG` at the top of `Common.gs`.
4. Run `setup()` and accept the permissions. Optional: run `loadSampleRequests()` for example data.
5. **Deploy > New deployment > Web app**. Execute as: *Me*. Who has access: *Anyone within hqaspring.org*.

- Teachers: the web app URL
- Principal: the web app URL + `?page=dashboard`

To update the code later, use **Deploy > Manage deployments > Edit > New version** so the link stays the same.

## Policy built in (Nov 6, 2026 staff meeting)

- Class parties: Thursdays only, within lunch & recess, planned as a team, with learning activities for the rest of the day.
- Volunteers and observers need approval in advance. Once approved, the Front Desk, the Principal and Ms. Omayya are emailed.

## Emails

| When | Who gets it |
|---|---|
| New or resubmitted request | Principal (cc `ADMIN_CC`); teacher gets a receipt |
| Approved / Needs Changes / Denied | Teacher, with the Principal's note; FYI copy to `ADMIN_CC` |
| Volunteer approved | `VOLUNTEER_NOTIFY` (Front Desk, Principal, Ms. Omayya) |
| Weekday mornings | Principal gets a summary of pending requests |

## Trial mode

`TEST_MODE: true` sends every email only to `TEST_EMAIL`, with a yellow bar showing who it would have gone to. `TEST_EMAIL` can also approve.

## Going live

- Set `TEST_MODE: false`.
- Fill in `APPROVERS`, `ADMIN_CC` and `VOLUNTEER_NOTIFY`.
- Set `PARTY_WINDOW_START` / `PARTY_WINDOW_END` to the real lunch & recess times.
- Delete any sample rows from the **Requests** tab.
- Own the Sheet and script from a school role account, and share the Sheet only with the people who need it.

## Access and security

- Data lives in a Google Sheet in HQA's Google Workspace.
- Only hqaspring.org accounts can open the app. Sign-in is Google's, so there are no extra passwords.
- Only `APPROVERS` can decide, and teachers see only their own requests. These checks run on Google's server, not in the browser.
- Every submission and decision is logged with who and when.
