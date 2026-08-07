# RA Workflow deployment

This build is based on the latest attached Claude version, with the agreed workflow implemented.

## Deploy

1. Back up the Supabase database.
2. Replace the repository files with this package.
3. Keep the existing Streamlit secrets; no new secret values are required.
4. Push to the branch deployed by Streamlit Community Cloud.
5. Open the app once as Manager. `init_db()` applies the additive schema migration automatically.

`schema.sql` is provided for reference or manual provisioning, but normal upgrades do not require running it separately.

## Important first-start behavior

- Existing `Closed` Workpage positions are migrated to `Exported`.
- The unfinished legacy `sessions` table, if present, is replaced. Users simply sign in again once.
- Prospect company lookup keys are changed to the exact stored company spelling. On a large Prospects DB, this one-time migration can make the first startup slower.
- Automatically assigned companies are now added to the Master (`companies`) table at assignment time.

## Verification checklist

- RA refreshes the browser and remains signed in.
- RA sees only Dashboard, Add company, Workpage, and EOD uploads.
- RA cannot access another RA's Workpage position.
- RA can edit Job Position, Job Location, JD URL, and saved prospect fields.
- Invalid/Bounced-DNC emails are blocked.
- Ready is blocked until every prospect is complete and the job position/location exist.
- Work Date persists between jobs and remains fixed on Ready records unless explicitly corrected.
- TL/Manager can reassign using active RA accounts.
- RA deletion requires the exact company-name confirmation.
- Not Eligible deletion creates a pending approval request and quarantines the company.
- Rejected requests and “Assigned by mistake” deletions appear in the reassignment queue.
- Campaign CSV columns are exactly `FIRST NAME`, `EMAIL ID`, `POSITION`, `LOCATION`.
- Export history can be re-downloaded and Exported positions require TL/Manager unlocking.

## Status lifecycle

`Open -> Ready -> Exported`

- Ready remains editable.
- Ready must return to Open before RA deletion.
- Exported is locked; TL/Manager can unlock it for correction or delete it with an immutable audit snapshot.
