# Geni Holocaust Projects — Monthly Profile Export

> **Repository note:** real credentials never belong in this repo. Copy
> `geni_config.example.json` to `geni_config.json` and fill in your own
> Geni app key/secret; set your own random key in `callback_index.php`
> (`$FETCH_KEY`) and the matching `auth_fetch_key` in the config. The
> `.gitignore` here keeps `geni_config.json`, token files, logs, and
> exports out of version control.

A server script that, once a month, walks the Geni umbrella project
[Holocaust: "The Final Solution"](https://www.geni.com/projects/Holocaust-The-Final-Solution/10996),
discovers all the sub-projects linked from it, downloads every profile in
every project through the Geni API, and writes **one combined spreadsheet
with a fixed, standardized set of columns**.

## Why the columns are always the same

Geni's own exports include a column for every field any profile happens to
use, and skip empty fields — so exports from different projects have
different shapes. This script sidesteps that entirely: instead of using
Geni's export, it fetches each profile as JSON from the API and maps
whatever fields come back onto a fixed column list. A profile with no
burial place simply gets an empty cell; the columns never change.

Output columns (in order): Profile ID, Profile URL, Name, First Name,
Middle Name, Last Name, Maiden Name, Gender, Birth Date, Birth Place,
Death Date, Death Place, Burial Place, Occupation, Projects, Profile
Created, Profile Updated.

Rows are alphabetized by surname, then given name — case-insensitively
and with accents folded (so Łucki sorts with Lucki). If a profile has no
Last Name, its Maiden Name is used for sorting; failing that, the last
word of the display name. Profiles with no usable name sort at the end.
A profile that belongs to several of the projects appears **once**, with
all its projects listed in the Projects column (separated by `;`). Each run
produces both a `.xlsx` (with a Run Info tab listing the projects covered)
and a `.csv`, date-stamped, in the `exports/` folder.

## Requirements

- Python 3.8+
- `pip install requests openpyxl`
- A Geni account (the script only reads data your account can see;
  profile privacy on Geni still applies through the API)

## One-time setup

### 1. Register a Geni application

1. Sign in to Geni, then go to <https://www.geni.com/platform/developer/apps>
   (Developer section → "Register a new application").
2. Give it any name (e.g. "Holocaust Projects Export"), and set the main
   URL / site domain to a domain you control — your server's domain is
   fine. The callback (redirect) URL must be on that same domain; it does
   not need to be a real working page. Example: `https://example.org/callback`.
3. After registering you get an **App Key** (`client_id`) and **App
   Secret** (`client_secret`).

### 2. Configure the script

Copy the files to your server, then run it once — it writes a template
`geni_config.json` next to itself:

```bash
python3 geni_holocaust_export.py
```

Edit `geni_config.json`:

```json
{
  "client_id":   "YOUR_APP_KEY",
  "client_secret": "YOUR_APP_SECRET",
  "redirect_uri": "https://example.org/callback",
  "umbrella_project_id": 10996,
  "discovery_depth": 1,
  "extra_project_ids": [],
  "exclude_project_ids": [],
  "include_umbrella_profiles": true,
  "output_dir": "exports",
  "token_file": "geni_tokens.json"
}
```

- `discovery_depth: 1` follows project links found in the umbrella
  project's description.
- Sub-projects that are themselves umbrellas are automatically expanded
  one more level: any discovered project whose name contains one of the
  `expand_name_patterns` (default: "umbrella", "portal") has the project
  links in its description followed too, down to `max_discovery_depth`.
  To force-expand a project whose name doesn't say "umbrella", add its ID
  to `expand_project_ids`.
- `extra_project_ids` / `exclude_project_ids`: add or remove specific
  projects by ID (the number at the end of a project's URL).

### 3. Authorize (one time)

```bash
python3 geni_holocaust_export.py --authorize
```

It prints a Geni URL. Open it in a browser, sign in and approve; the
browser then lands on your redirect URL with `?code=...` in the address
bar (the page itself may 404 — that's fine, you only need the address
bar). Paste the code back into the script. Tokens are saved to
`geni_tokens.json` and refreshed automatically on every future run.

### 4. Sanity checks

```bash
# Which projects would be exported?
python3 geni_holocaust_export.py --projects-only

# Small trial run: at most 100 profiles per project
python3 geni_holocaust_export.py --max-profiles 100
```

Review the project list before the first full run — the script discovers
sub-projects from the links in the umbrella project's description, so
`--projects-only` shows you exactly what it found, and you can prune with
`exclude_project_ids` or pin additions with `extra_project_ids`.

## Monthly schedule (hourly watchdog cron)

The script keeps a checkpoint of its progress (`exports/run_state.json`
plus the `partial_run.csv` snapshot, updated after every project and
every 20 pages inside big projects). The recommended schedule is an
HOURLY cron entry in `--cron` watchdog mode:

```cron
0 * * * * cd /path/to/geni_export && python3 geni_holocaust_export.py --cron >> cron.log 2>&1
```

Each hour it checks: is this month's export already finished? (exit) Is
a run currently in progress? (exit) Otherwise it starts — or resumes —
the month's export from the last checkpoint. The result: the export
starts automatically at the top of the month, survives kills, reboots,
and closed terminals by resuming where it stopped, and once the month's
spreadsheet is written the hourly checks cost nothing. Never start long
runs from a web-hosted terminal (cPanel/WHM Terminal) — those kill
everything they started the moment the tab closes; let the cron
watchdog do the starting.

Each completed run leaves a date-stamped `.xlsx` and `.csv` in
`exports/`, emails you download links, and appends progress to
`geni_export.log`.

## Re-authorization by email

Geni's stored refresh token normally keeps working from month to month,
but if it ever stops (revoked, expired, security event on the account),
the script does not just fail. Before each run it verifies authorization,
and if authorization breaks — before or in the middle of a run — it:

1. Emails `notify_email` a Geni authorization link.
2. Waits, checking every minute (up to `auth_wait_hours`), while you:
   open the link, approve, copy the **entire address** of the page you
   land on (it contains `?code=...`), and save it into a plain text file
   at the location named in the email (`geni_auth_code.txt` next to the
   script by default).
3. Picks up the file, exchanges the code, and **resumes the run exactly
   where it stopped** — no downloaded data is lost.

### Fully automatic: one click, no copy-paste

Because the registered Callback URL is on a site you control, step 2 can
be automated away. Upload `callback_index.php` to the schoenberg.com
webroot as `callback/index.php`, so it answers at
`https://schoenberg.com/callback`. Then when authorization is needed:
you click the link in the email, approve on Geni, and see an
"Authorization received" page — that's the whole procedure. The callback
page catches the code and the script (which polls it every minute during
the wait, via `auth_code_url` + `auth_fetch_key` in the config) picks it
up and resumes on its own. The `auth_fetch_key` in the config must match
`$FETCH_KEY` inside the PHP file (already set to the same value in both).
The manual text-file fallback below still works if the PHP page isn't
deployed.

Tip: if the script folder on the server is synced with Dropbox, you can
create `geni_auth_code.txt` from any of your devices.

Notification emails are sent the same way the Virtual Arnold daily
digest sends its emails: through the web server's own mail() — no SMTP
credentials needed. The script POSTs the message to the callback page
(`callback_index.php`), which mails it to `notify_email`
(randols912@gmail.com) from geni-export@schoenberg.com. This works as
soon as the callback page is deployed. Alternatives via `mail_method` in
the config: `sendmail` (when the script runs directly on the web host)
or `smtp` (fill in the `smtp` block). If email can't be sent, the same
instructions are written to `geni_export.log` instead.

The authorization code expires ~10 minutes after you approve, so save the
file promptly; if it expires, the script emails you again and keeps
waiting.

## Notes & troubleshooting

- **Rate limits.** Geni's initial API limit is 40 requests per 10
  seconds. The script reads Geni's rate-limit headers and slows itself
  down, and backs off automatically on HTTP 429 — a large export just
  takes longer, it doesn't fail. At 50 profiles per request, ~100,000
  profiles is roughly 2,000 requests (well under an hour).
- **"No saved tokens" / refresh fails.** If the export hasn't run for a
  long time or you revoked the app, run `--authorize` again.
- **New sub-projects.** Because discovery re-reads the umbrella project
  each run, newly added sub-projects are picked up automatically the next
  month.
- **Testing without the network.** `python3 test_normalize.py` runs the
  whole pipeline against canned API responses.
- **Privacy safeguard:** with `public_only` (on by default), the script
  exports only profiles marked public on Geni. Profiles the authorizing
  account can see through family or curator privileges but which are not
  public are skipped (the log reports how many). For extra assurance,
  authorize the app with a dedicated non-curator Geni account rather than
  a curator account — then non-public data is never even visible to the
  script.
