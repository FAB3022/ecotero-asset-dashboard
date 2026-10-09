# Asset securement dashboard

A dashboard for a brand asset tracker kept in Google Sheets. It shows where each asset stands against an evidence gate, and it edits the sheet directly, so the sheet and the dashboard always show the same thing.

This repository holds code only. It contains no project data, no sheet links and no keys.

## How it works

```
Browser (this page)  <-->  Apps Script web app  <-->  Tracker sheet ("Asset Register" tab)
                                                 <-->  Project Log sheet (plan, evidence, weekly updates)
```

- The **tracker sheet** is the single source for assets. The dashboard reads it about every 45 seconds and writes each edit straight back to the row with the matching Asset ID.
- Columns are found by their header text, never by position, so columns can be added, moved or removed in the sheet. Dropdown choices are read from the sheet's own rules.
- The **Project Log** is a second spreadsheet the script creates in the owner's Drive. It holds milestones, actions, evidence, decisions, team, weekly updates, weekly snapshots, a change log and removed rows. The tracker itself is left with its own tabs.
- Every Friday the script records a snapshot and writes a draft weekly update with week-over-week changes. It sends nothing.
- Nothing loads until a visitor enters the team **access key**. The key is stored in that browser only.

## Set up the Google side (about five minutes, done once by the sheet owner)

1. Open the tracker sheet, then **Extensions > Apps Script**.
2. Delete what is in the editor, paste the contents of [`apps-script/Code.gs`](apps-script/Code.gs), and save.
3. Choose the function **setup** in the toolbar and press **Run**. Approve the permission prompt.
4. Open **Execution log**. Copy the **access key** and note the Project Log link.
5. **Deploy > New deployment**, type **Web app**. Set *Execute as* to **Me** and *Who has access* to **Anyone**. Deploy and copy the **web app URL**.
6. Open the dashboard, paste the web app URL and the access key, and connect.

To give the team one-step access, put the web app URL in [`config.js`](config.js). They then only enter the key.

After changing `Code.gs` later: **Deploy > Manage deployments > Edit > Version: New version > Deploy**. The URL stays the same.

## Security notes

- The web app runs as the sheet owner, so the sheet does not need to be shared by link for the dashboard to work. Restrict the sheet to named people.
- Share the access key privately. To replace it, run **rotateKey** in the script editor.
- The dashboard refuses to save values that look like passwords or recovery codes. Keep those in a password manager and record only the vault reference.
- An asset can be saved as "Secured - verified" from the dashboard only when every gate check that applies passes. Edits made directly in the sheet are not blocked, so the dashboard flags any asset marked verified with checks still open.

## Files

| File | Purpose |
| --- | --- |
| `index.html` | The dashboard. One file, no build step. |
| `config.js` | Optional default web app URL. |
| `apps-script/Code.gs` | The script that sits between the dashboard and the sheets. |
