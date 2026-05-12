# Dashboard auto-refresh setup

The dashboard rebuilds itself every 6 hours from the Google Sheet at
`https://docs.google.com/spreadsheets/d/1v2b7OMQ3Hvz9LwMroYw23CHIn3NsBZhFxDezy-QzcXQ/`.

## How it works

1. `.github/workflows/refresh.yml` cron-triggers `scripts/build_dashboard.py`.
2. The script reads the Sheet via a Google service account, computes all stats, and
   renders `scripts/template.html` into `index.html`.
3. If `index.html` changed, the workflow commits and pushes — Pages redeploys.

## One-time setup (do this once)

### 1. Create a Google service account

1. Go to https://console.cloud.google.com → create a new project (free).
2. Search for "Google Sheets API" → enable.
3. **Credentials** → **Create credentials** → **Service account**. Name it
   `pof-dashboard-reader`. Skip the role-grant step.
4. Open the new service account → **Keys** tab → **Add key** → **JSON**. A `.json`
   file downloads — keep it safe, you'll paste it in step 3.

### 2. Share the Sheet with the service account

The downloaded JSON has a field like `"client_email":
"pof-dashboard-reader@<project>.iam.gserviceaccount.com"`. Copy that email.

Open the POF Leader LIST sheet → **Share** → paste the email → **Viewer**
permission → **Send**.

### 3. Add the JSON to GitHub secrets

In the `lizmckenna/pof-dashboard` repo:
**Settings** → **Secrets and variables** → **Actions** → **New repository secret**:

- Name: `GOOGLE_SERVICE_ACCOUNT_JSON`
- Value: paste the entire contents of the JSON file from step 1

### 4. Trigger the first build

**Actions** tab → **Refresh dashboard** → **Run workflow**.

Wait ~1 minute, then check https://lizmckenna.github.io/pof-dashboard/ — the
data should now reflect the live Sheet.

## Local development

```bash
cd /tmp/pof-dashboard
pip install -r scripts/requirements.txt
python scripts/build_dashboard.py
```

By default, this reads the local xlsx at `~/Desktop/POF Leader LIST.xlsx` (or
`~/Desktop/eb/POF Leader LIST.xlsx`).

To test the Sheets path locally, export the JSON inline:
```bash
export GOOGLE_SERVICE_ACCOUNT_JSON="$(cat /path/to/key.json)"
python scripts/build_dashboard.py
```

## When something breaks

- **"Warnings: <fellow>: no First Name column in ..."** — a fellow renamed a
  column. Either fix the column name in their tab, or add the new name to
  `HEADER_ALIASES` in `scripts/build_dashboard.py`.
- **"Module not found: gspread"** — run `pip install -r scripts/requirements.txt`.
- **GitHub Action fails with 403** — the service account doesn't have access to
  the Sheet. Re-share the Sheet with the service account email.
- **No changes appearing** — Pages can take ~1 minute after commit. Check Actions
  tab for the latest run's status.

## Header drift — what to tell fellows

> The dashboard pulls your numbers automatically from your tracker tabs. Please
> keep these column headers named exactly as the template has them: **First
> Name, Last Name, Date of 1-1, Notes from 1-1, What is their self-interest?,
> Invitation Status, How did you find them?** You can add your own columns —
> just don't rename the ones already there. If the dashboard looks off, that's
> usually why.
