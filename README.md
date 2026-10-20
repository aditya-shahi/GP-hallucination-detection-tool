# Domain-wise Sheet — Hallucination Detection Tool

A local, single-file Flask app that checks whether the values in a
**domain-wise Excel sheet** actually appear on the pages they claim to
come from.

Two pages:

1. **Checker** — reads a Fund `.xlsx` (sheet `Sheet1`) and, for every row
   and every datapoint, verifies that the cell value appears in the visible
   text of the reference URL(s) listed next to it.
2. **Keyword Match** — reads a URL-wise CSV, filters its `Domain` column by
   the domain(s) you type, and searches every page under those domains for a
   per-domain keyword list.

Everything runs on your machine. The only outbound calls are:

- `requests` → the internal **forage decom** endpoint (for `s3status` links),
- `gspread` → Google Sheets / Drive **only when you choose Google Sheet
  output**.

---

## 1. Install

```bash
python -m venv .venv
source .venv/bin/activate           # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

---

## 2. Run

```bash
python frontend.py
```

Open <http://localhost:8767>. The port is `8767` (see `PORT` in `frontend.py`).

Default login for the internal forage server can be supplied on each page
or set as environment variables (see §6).

---

## 3. Inputs

### Fund Excel (Checker page)

- `.xlsx` file with a sheet named **`Sheet1`** (falls back to the first
  sheet if `Sheet1` is absent).
- Row 1 must contain the column headers.
- Recognised datapoints and their reference-URL columns:

  | Family | Value column | Reference-URL column |
  |---|---|---|
  | AUM | `Firm AUM` | `Firm AUM Reference URL` |
  | Date | `Firm AUM Date` | `Firm AUM Date Reference URL` |
  | Plain list | `Firm Geo Focus` | `Firm Geo Focus Reference URL` |
  | Plain list | `Firm Country Focus` | `Firm Country Focus Reference URL` |
  | Plain list | `Firm Industry Focus` | `Firm Industry Focus Reference URL` |
  | Asset-class list | `Asset Class Geo Focus` | `Asset Class Geo Focus Reference URL` |
  | Asset-class list | `Asset Class Country Focus` | `Asset Class Country Focus Reference URL` |
  | Asset-class list | `Asset Class Industry Focus` | `Asset Class Industry Focus Reference URL` |
  | Asset-class list | `Investment Type` | `Investment Type Reference URL` |
  | Asset-class date | `Asset Class AUM Date` | `Asset Class AUM Date Reference URL` |
  | Asset-class numbers | `Asset Class AUM` | `Asset Class AUM Reference URL` |
  | Asset-class numbers | `Total Capital Invested` | `Total Capital Invested Reference URL` |

- **Ignored (no sheet, no check):**
  `Firm AUM Currency`, `Asset Class`, `Total Capital Invested Currency`.
- A reference-URL cell may contain several URLs separated by `|`.

### Value formats

| Family | Cell format | Notes |
|---|---|---|
| AUM | `1,234.5` | Number extracted by regex; matches `1,234.5`, `1234.5`, `1234.50` |
| Date | `DD/MM/YYYY` | Also matched on the page as `DD-MM-YYYY`, `MM/DD/YYYY`, `DD/MM/YY`, `MM/DD/YY`, unpadded |
| Plain list | `India, China, Indonesia \| Mexico \| France` | Split on `,` **and** `\|`; deduped case-insensitively |
| Asset-class list | `Equity - India, China \| Private Debt - Mexico` | Split on `\|`, drop the `Class - ` prefix, split remainder on `,` |
| Asset-class date | `Equity - 12/03/2020 \| Private Debt - 05/08/2019, 22/11/2021` | Strip prefix first, each `Data` is `DD/MM/YYYY` |
| Asset-class numbers | `Equity - 1,250 \| Private Debt - 800, 450` | Strip prefix first, each `Data` is a number |

**Rule (B)** applies to every list / asset-list / asset-date / asset-numbers
datapoint: a row is a hit **only if every listed value is found somewhere
across the reference URLs**. The result sheet always reports both what was
found and what was missing, so rule (A) ("at least one value found") is
recoverable by filtering the output.

### URL-wise CSV (both pages)

Required columns (header row, case-sensitive):

| Column | Meaning |
|---|---|
| `URL` | Page URL (used as a human-readable label in outputs) |
| `s3status` | Link to the page's JSON payload; the app fetches this and extracts `url_visible_text` |
| `Domain` | Domain this row belongs to (Checker page ignores this column; Keyword Match uses it) |

The Checker page only needs `URL` and `s3status`; the Keyword Match page
additionally requires `Domain`.

---

## 4. Using the Checker page

1. **Fund Excel (.xlsx)** — the file described in §3.
2. **URL-wise CSV** — as above.
3. **Login ID / Password** — credentials for the forage decom endpoint
   (can be left blank if set via env vars, see §6).
4. **Output** — Google Sheet or Download .xlsx.
5. **Run check**.

For each row and each of the 12 datapoints, the app:

- looks up each reference URL in the URL-wise CSV (normalised, `www.` kept),
- fetches the `s3status` link, extracts `url_visible_text`,
- runs the appropriate matcher,
- writes a row of results.

You'll get **12 output sheets**, one per datapoint, named as below:

| Sheet | Contents |
|---|---|
| `AUM_results` | Firm AUM match columns + `Per-URL Results` |
| `Date_results` | Firm AUM Date match columns + `Per-URL Results` |
| `Firm_Geo_Focus_results` | Plain list — `Values searched / count / found / missing`, `All values found`, `Per-value evidence`, `Per-URL Results` |
| `Firm_Country_Focus_results` | …same shape… |
| `Firm_Industry_Focus_results` | …same shape… |
| `Asset_Class_Geo_Focus_results` | Asset-class list — `Class + data searched` plus the plain-list columns |
| `Asset_Class_Country_Focus_results` | …same shape… |
| `Asset_Class_Industry_Focus_results` | …same shape… |
| `Investment_Type_results` | …same shape… |
| `Asset_Class_AUM_Date_results` | Asset-class date — `Dates searched / count / parsed / unparseable / found / missing`, `All dates found`, `Matched version`, `Per-date evidence` |
| `Asset_Class_AUM_results` | Asset-class numbers — `Numbers searched / count / found / missing`, `All numbers found`, `Per-number evidence` |
| `Total_Capital_Invested_results` | …same shape… |

The progress panel shows:

- **Processed** (rows done / total rows),
- **AUM** / **Date** count of rows that matched,
- **plain / asset-list / asset-date / asset-AUM pass** counts (rows where
  *every* value was found).

---

## 5. Using the Keyword Match page

1. **Input mode**
   - *Google Sheet (remembered)* — reuse the sheet from the last run.
   - *Download .xlsx* — get a standalone workbook.
2. **Domains & keywords** — one or more lines. Format:

   ```
   <domain>[, <domain> ...]: "<kw>", "<kw>", ...
   ```

   Examples:

   ```
   eatonvance.com: "AUM", "assets"
   morganstanley.com: "fee-based", "advisory"
   morganstanley.com, blackrock.com: "AUM", "assets"
   ```

   - Several domains before the colon share the same keyword list.
   - Duplicate domains across lines are **unioned** (case-insensitive).
   - The domain must match a value in the URL-wise CSV's `Domain` column
     **exactly** (case-insensitive, surrounding whitespace trimmed).
   - Keywords are matched as whole words, case-insensitive,
     whitespace-flexible.

3. **Login ID / Password** — same as the Checker page.
4. **URL-wise CSV (optional)** — upload to replace the remembered copy.
5. **Match**.

For each input domain, the app:

- filters the URL-wise rows where `Domain == <input>`,
- fetches each row's `s3status` → `url_visible_text`,
- matches every keyword for that domain against every page.

Output: **one row per domain** in a single sheet named
**`Domain_keyword_matches`**:

| Column | Meaning |
|---|---|
| `Domain` | Input domain (normalised) |
| `Keywords` | All keywords assigned to that domain, `\|`-joined |
| `URLs scanned` | URL-wise rows with this Domain and a non-blank s3status |
| `Matched with` | Keywords found on at least one page, `\|`-joined |
| `URLs matched` | `URL` values where at least one keyword was found, `\|`-joined |
| `Evidence` | `URL :: context` lines (10 words each side, match in `[]`), `\|`-joined |
| `Notes` | Fetch errors, blank s3status, `"domain not in URL-wise file"`, etc. |

A domain that is not present in the URL-wise CSV still gets a row, with
`URLs scanned = 0` and `Notes = "domain not in URL-wise file"` — so typos
are visible in the output.

---

## 6. Configuration (environment variables)

All are optional; the page fields override them when filled.

| Variable | Default | Purpose |
|---|---|---|
| `GOOGLE_SA_JSON` | `service_account.json` | Path to the Google service-account JSON key |
| `GOOGLE_SHEET_ID` | *(empty)* | A default Google Sheet ID / URL to write to (leave empty to create a new sheet each run) |
| `S3_LOGIN_ID` | *(empty)* | Default login for the forage decom endpoint |
| `S3_PASSWORD` | *(empty)* | Default password for the forage decom endpoint |
| `CHECKER_STATE_FILE` | `checker_state.json` | Where the "remembered" Google Sheet + URL-wise data is stored |

Example:

```bash
export S3_LOGIN_ID="my-user"
export S3_PASSWORD="my-pass"
export GOOGLE_SA_JSON="/etc/secrets/service_account.json"
python frontend.py
```

---

## 7. Google Sheets setup (optional, free)

Only needed if you choose **Google Sheet** output.

1. <https://console.cloud.google.com> → create a project.
2. **APIs & Services → Library** → enable **Google Sheets API**.
3. **APIs & Services → Library** → enable **Google Drive API**
   (needed to create new sheets).
4. **IAM & Admin → Service Accounts** → create a service account.
5. Open it → **Keys → Add key → Create new key → JSON**.
   Save the downloaded file next to `frontend.py` as
   **`service_account.json`** (or point `GOOGLE_SA_JSON` at it).
6. On the Checker page (or the Keyword Match page), either:
   - **leave the Google Sheet URL blank** and type your own Google email in
     the *Your Google email* field — a new sheet is created and shared with
     you as **Editor**; or
   - **paste an existing sheet's URL** and share that sheet with the
     service account's `client_email` (visible at `/health`) as **Editor**.

`/health` reports whether the service-account file is present and what
its `client_email` is:

```bash
curl -s http://localhost:8767/health | python -m json.tool
```

> If the service account has no Drive storage and you ask the app to
> *create* a sheet, creation will fail with an HTTP error. In that case,
> create the sheet yourself and paste its URL.

---

## 8. Where "remembered" data lives

`checker_state.json` (next to the script by default, or the path in
`CHECKER_STATE_FILE`) stores:

- the last Google Sheet ID / URL / title,
- the last URL-wise data (`lookup` map after a Checker run,
  `lookup_rows` after a Keyword Match run),
- timestamps.

Passwords are **never** written to this file.

To reset:

```bash
rm checker_state.json
```

---

## 9. Troubleshooting

| Symptom | Likely cause / fix |
|---|---|
| `could not read the Fund .xlsx` | The file is not `.xlsx`, or `Sheet1` is empty. Rename the sheet or put headers in row 1. |
| `URL-wise CSV is missing required column 'Domain'` | The Keyword Match page needs the `Domain` column; the Checker page doesn't. |
| `No URL-wise data saved yet` (Keyword Match) | Upload the URL-wise CSV once, or run the Checker first. |
| Every row shows `URL not in URL-wise file` | The reference URLs in the Fund sheet don't match (after normalisation) the `URL` column of the URL-wise CSV. Both are normalised the same way: lowercase host, trailing `/` trimmed, query kept. |
| `Service-account key not found at …` | Set `GOOGLE_SA_JSON` or drop `service_account.json` next to `frontend.py`. Or use **Download .xlsx**. |
| `Cannot open the Google Sheet (HTTP 403)` | Share the sheet with the service account's `client_email` as Editor. |
| `The sheet was created … but sharing it … failed` | Your Google account can't be reached by the service account. Create the sheet yourself and share it out-of-band. |
| Google rate-limit warnings | The app retries with backoff (15/30/45/60 s). Large sheets may take a few minutes. |
| `domain not in URL-wise file` | Exact-string mismatch with the `Domain` column. Only surrounding whitespace and case are ignored. |

---

## 10. Project layout

```
.
├── frontend.py            # the whole app (backend + frontend)
├── requirements.txt       # Python dependencies
├── README.md              # this file
├── service_account.json   # optional: Google service-account key
└── checker_state.json     # created at first run — remembered sheet + URL-wise data
```

- `frontend.py` is the whole app: Flask backend, all matchers, both Google
  Sheets and Excel writers, plus the embedded HTML/JS front-end
  (`PAGE_HTML`). There are **no other source files**.
- The app writes **12 sheets** from the Checker page and **1 sheet** from
  the Keyword Match page.

---

## 11. HTTP endpoints (for debugging)

| Method | Path | Purpose |
|---|---|---|
| `GET` | `/` | The UI |
| `GET` | `/health` | Version/state probe: gsheet readiness, service-account e-mail, default sheet |
| `GET` | `/api/state` | Remembered Google Sheet + URL-wise data summary |
| `POST` | `/api/run-stream` | Start a Checker run; returns `{stream_id}` |
| `POST` | `/api/domain-match-stream` | Start a Keyword Match run; returns `{stream_id}` |
| `GET` | `/api/progress/<sid>` | SSE stream of progress events for a run |

SSE event kinds: `progress`, `status`, `done`, `fail`, `ping`, `end`.

---

## 12. Notes on matching

- **AUM**: matches `1,234.5`, `1234.5`, `1234.50`, with word boundaries so
  `11234` won't match `1234`. Comma and decimal-padding variants are all
  covered.
- **Date**: tries all six of `DD/MM/YYYY`, `DD-MM-YYYY`, `MM/DD/YYYY`,
  `DD/MM/YY`, `MM/DD/YY`, and unpadded `D/M/YYYY`. The result sheet records
  which version(s) matched.
- **List / asset-list / asset-date / asset-numbers**: whole-word,
  case-insensitive, whitespace-flexible literal match. A value like
  `Real Assets` will match `real  assets` on the page but not
  `unrealassets`.
- **Every** match is evidence-tagged with 10 words of context on each side;
  the matched span is wrapped in `[…]` so you can see exactly what hit.
- **Rule (B)** (all values must be found) is enforced everywhere for the
  list-like families; per-value hit/miss is always recorded so you can
  apply rule (A) downstream by filtering.

---

## 13. Licence

Internal tool. No warranty; treat outputs as evidence, not ground truth.