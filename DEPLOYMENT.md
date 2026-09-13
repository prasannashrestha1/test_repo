# Quiet List Exchange Activity Report — VPS deployment runbook

Generates the Quiet List Exchange Activity Report (one PDF per office) on the
1st and 15th of every month, pulling live data from Bubble's Data API.

Runs as a **systemd timer** on a Linux VPS. There is no long-running process
and no container — the timer wakes a one-shot job, it generates the PDFs, and
it exits.

Tested against Ubuntu 22.04/24.04. Debian 12 works the same way.

---

## 1. Prerequisites

- A Linux VPS you control, with root/sudo.
- Outbound HTTPS to `app.quietlist.com.au` (the Bubble Data API).
- A Bubble API token with read access to the `Property_Matches` and
  `property` objects.

## 2. Create the service user and directory

A dedicated unprivileged user — the job never needs root.

```bash
sudo useradd --system --create-home --home-dir /opt/quietlist-reports --shell /usr/sbin/nologin quietlist
```

## 3. Copy the application onto the box

From your machine (**note the trailing `/` and the excludes** — never ship
your own `.env`, and don't bother shipping local build junk):

```bash
rsync -av --exclude='.env' --exclude='venv' --exclude='__pycache__' \
      --exclude='reports' --exclude='logs' \
      ./ root@YOUR_VPS:/opt/quietlist-reports/
```

Then on the VPS:

```bash
sudo mkdir -p /opt/quietlist-reports/reports /opt/quietlist-reports/logs
sudo chown -R quietlist:quietlist /opt/quietlist-reports
```

## 4. Install Python dependencies and Chromium

```bash
sudo apt update
sudo apt install -y python3 python3-venv python3-pip

cd /opt/quietlist-reports
sudo -u quietlist python3 -m venv venv
sudo -u quietlist venv/bin/pip install -r requirements.txt

# Installs Chromium AND the system libraries it needs. The apt step needs
# root, which is why this one is not run as the quietlist user.
sudo venv/bin/playwright install --with-deps chromium

# Chromium landed in root's cache above; put it where the service user reads it.
sudo mkdir -p /opt/quietlist-reports/.cache
sudo cp -r /root/.cache/ms-playwright /opt/quietlist-reports/.cache/
sudo chown -R quietlist:quietlist /opt/quietlist-reports/.cache
```

> If Chromium can't be found at runtime, that last copy is almost always why.
> Playwright looks in `$HOME/.cache/ms-playwright`, and `$HOME` for the
> service is `/opt/quietlist-reports`.

## 5. Supply the API token

```bash
sudo -u quietlist cp /opt/quietlist-reports/.env.example /opt/quietlist-reports/.env
sudo -u quietlist nano /opt/quietlist-reports/.env      # paste the real token
sudo chmod 600 /opt/quietlist-reports/.env
```

`chmod 600` matters — it keeps the token readable only by the service user.

## 6. Enable production access

**The pipeline ships with production deliberately disabled.** `config.json`
has no `bubble_base_url`, so a real run fails fast instead of reaching live
data (see `PRODUCTION_ACCESS_DISABLED.md`). Until you add it back, the timer
will fire and log an error rather than produce reports.

Add this line to `config.json`, next to `bubble_base_url_test`:

```json
"bubble_base_url": "https://app.quietlist.com.au/api/1.1/obj",
```

## 7. Install the systemd units

```bash
sudo cp /opt/quietlist-reports/deploy/quietlist-report.service /etc/systemd/system/
sudo cp /opt/quietlist-reports/deploy/quietlist-report.timer   /etc/systemd/system/
sudo systemctl daemon-reload
sudo systemctl enable --now quietlist-report.timer
```

## 8. Verify

```bash
# Confirm the next scheduled run looks right (should be the next 1st or 15th)
systemctl list-timers quietlist-report.timer

# Force a run right now, without waiting for the timer
sudo systemctl start quietlist-report.service

# Watch what it did
journalctl -u quietlist-report.service -n 50 --no-pager
```

A successful run logs a line like `1 offices, 1 reports, 0 errors` and drops
the PDF in `/opt/quietlist-reports/reports/`.

On any day that isn't the 1st or 15th, a forced run correctly logs
`Not a reporting day (today is the N); skipping.` and exits 0 — that's the
date guard doing its job, not a failure.

---

## Operations

| Task | Command |
|---|---|
| Next scheduled run | `systemctl list-timers quietlist-report.timer` |
| Logs (last run) | `journalctl -u quietlist-report.service -n 50 --no-pager` |
| Logs (follow live) | `journalctl -u quietlist-report.service -f` |
| Run now, off-schedule | `sudo systemctl start quietlist-report.service` |
| Pause the schedule | `sudo systemctl disable --now quietlist-report.timer` |
| Resume the schedule | `sudo systemctl enable --now quietlist-report.timer` |
| Change the schedule | edit `OnCalendar=` in the timer, then `sudo systemctl daemon-reload && sudo systemctl restart quietlist-report.timer` |

Generated PDFs: `/opt/quietlist-reports/reports/`

### Changing which offices get reports

Edit `offices.json`, then just wait for the next run — no restart needed,
since each run re-reads it from disk.

### Manual one-off run for a specific period

```bash
cd /opt/quietlist-reports
sudo -u quietlist venv/bin/python generate_report.py \
    --use-test-version --only-office 1 \
    --period-start 2026-08-19 --period-end 2026-09-06 \
    --prev-period-start 2026-07-01 --prev-period-end 2026-07-31
```

Drop `--use-test-version` to run against production (requires step 6).
Add `--mock` to test the PDF pipeline with synthetic data and no API calls
at all.

---

## Troubleshooting

| Symptom | Cause |
|---|---|
| `config.json is missing "bubble_base_url"` | Step 6 not done. This is the intended fail-safe, not a bug. |
| `BUBBLE_API_TOKEN environment variable is not set` | `.env` missing/empty, or not readable by the `quietlist` user. |
| `Executable doesn't exist at .../ms-playwright/...` | The Chromium cache copy in step 4 was skipped. |
| `404 Not Found` on an object URL | A Bubble object name in `config.json` doesn't match the app, or that data type isn't exposed via Bubble's Data API. |
| Timer never fires | `systemctl is-enabled quietlist-report.timer` — enable it, and check the clock/timezone with `timedatectl`. |

### Upgrading

```bash
rsync -av --exclude='.env' --exclude='venv' --exclude='__pycache__' \
      --exclude='reports' --exclude='logs' \
      ./ root@YOUR_VPS:/opt/quietlist-reports/
sudo chown -R quietlist:quietlist /opt/quietlist-reports
sudo -u quietlist /opt/quietlist-reports/venv/bin/pip install -r /opt/quietlist-reports/requirements.txt
```

No restart needed — the next timer firing picks up the new code. Only re-copy
the systemd units (step 7) if the files under `deploy/` changed.
