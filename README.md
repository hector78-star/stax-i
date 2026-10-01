# STAX-I sign-up: Motion 001

A one-page vote and sign-up form for the open borders debate (Thu 15 Oct 2026, 7pm, Buchanan Theatre).

- `docs/` is the static page that the QR code points to. With `API_URL` empty it runs in demo mode, saving to the current browser only.
- `apps-script/Code.gs` is the backend. It runs on a Google Sheet.

## Connect the Google Sheet (about 5 minutes)

1. Create a new Google Sheet called "STAX-I Motion 001".
2. Open Extensions → Apps Script, delete the starter code, paste in `apps-script/Code.gs`, and save.
3. Pick `setup` from the function dropdown and click Run. Approve the permissions prompt. This creates the Tally and Signups tabs.
4. Click Deploy → New deployment → type: Web app. Set Execute as to **Me** and Who has access to **Anyone**, then Deploy.
5. Copy the web app URL (it ends in `/exec`) and paste it into `API_URL` near the bottom of `docs/index.html`.

If you change `Code.gs` later, use Deploy → Manage deployments → Edit → New version. That keeps the same URL.

## Privacy design

- Votes are never stored as rows. Each vote adds one to a counter in the Tally sheet, so no vote can be linked to a person.

- The "places taken" count is `SEED_PLACES` (32) plus the number of "yes" sign-ups. The vote count starts at 22 yes / 14 no (`SEED_VOTES`, set in both files).
