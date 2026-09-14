"""email_report.py — emails generated report files via the SendGrid Web API.

Switched from Gmail SMTP: a Claude Code cloud routine's network sandbox is a
domain allowlist over HTTPS/443 only, not general TCP, so SMTP's port 587
was unreachable from the routine no matter what (smtp.gmail.com itself was
reachable, TLS/auth all worked locally) -- this is a sandbox limitation, not
a bug in the SMTP code. SendGrid's Mail Send endpoint is a plain HTTPS POST,
so the exact same call works both locally and from the routine once its
domain is allowlisted.

No SMTP, no smtplib, no new dependency -- `requests` (already required by
generate_report.py for the Bubble API) posts directly to
https://api.sendgrid.com/v3/mail/send with the API key as a Bearer token.

One-time setup:
  1. Create a SendGrid account (sendgrid.com) -- the free tier (100
     emails/day) covers this use case.
  2. Verify a sender: Settings -> Sender Authentication -> either verify a
     single email address ("Single Sender Verification" -- fastest, just
     click a confirmation link sent to it) or authenticate a whole domain
     (better deliverability, needs a few DNS records added at your domain
     registrar). Whichever address ends up verified is what
     SENDGRID_FROM_EMAIL below must be set to -- SendGrid silently rejects
     sends from an unverified address.
  3. Create an API key: Settings -> API Keys -> Create API Key. Restricted
     Access with only "Mail Send" enabled is enough; Full Access isn't needed.
  4. Set these as environment variables / secrets (never in config.json,
     never committed):
       SENDGRID_API_KEY   - the key from step 3 (a secret)
       SENDGRID_FROM_EMAIL - the verified address from step 2 (not secret,
                             but still an env var so it isn't hard-coded)
  5. If running as a Claude Code routine, allowlist api.sendgrid.com on the
     cloud environment's network access settings -- the same "Allowed
     domains" list app.quietlist.com.au is already on.

Sending is a deliberate no-op, not an error, when either variable is unset,
so mock/test-version runs and local development never need email access at
all.
"""
import base64
import mimetypes
import os

import requests

SENDGRID_ENDPOINT = "https://api.sendgrid.com/v3/mail/send"


def send_report_email(paths: list, to_addrs: list, subject: str = None, body: str = None) -> dict:
    """Email every file in paths as attachments, in one message, to to_addrs.
    Works for any file type -- PDFs and the .docx report alike.

    Returns {"status": ..., ...}:
      "sent"    — delivered; includes "to" and "count"
      "skipped" — not an error; includes "reason" (no files/recipients/credentials)
      "failed"  — includes "error"
    """
    if not paths:
        return {"status": "skipped", "reason": "no files to send"}

    if not to_addrs:
        print("no recipients configured — skipping email delivery")
        return {"status": "skipped", "reason": "no recipients"}

    # .strip() because a key pasted into a secrets UI routinely picks up a
    # trailing newline, which an Authorization header then sends literally.
    api_key = (os.environ.get("SENDGRID_API_KEY") or "").strip()
    from_email = (os.environ.get("SENDGRID_FROM_EMAIL") or "").strip()
    if not api_key or not from_email:
        print("SENDGRID_API_KEY/SENDGRID_FROM_EMAIL not set — skipping email delivery")
        return {"status": "skipped", "reason": "no SendGrid credentials configured"}

    try:
        attachments = []
        for path in paths:
            mime_type, _ = mimetypes.guess_type(path)
            with open(path, "rb") as f:
                content_b64 = base64.b64encode(f.read()).decode("ascii")
            attachments.append({
                "content": content_b64,
                "filename": os.path.basename(path),
                "type": mime_type or "application/octet-stream",
                "disposition": "attachment",
            })

        payload = {
            "personalizations": [{"to": [{"email": addr} for addr in to_addrs]}],
            "from": {"email": from_email},
            "subject": subject or "Quiet List Exchange Activity Report",
            "content": [{
                "type": "text/plain",
                "value": body or "Attached: the latest Quiet List Exchange Activity Report(s).",
            }],
            "attachments": attachments,
        }

        response = requests.post(
            SENDGRID_ENDPOINT,
            headers={"Authorization": f"Bearer {api_key}"},
            json=payload,
            timeout=30,
        )
        # SendGrid returns 202 with an empty body on success; on failure the
        # real reason (bad key, unverified sender, etc.) is in the response
        # body, which is far more useful to surface than a bare status code.
        if response.status_code != 202:
            raise RuntimeError(f"SendGrid returned {response.status_code}: {response.text}")

        print(f"emailed {len(paths)} report file(s) to {', '.join(to_addrs)}")
        return {"status": "sent", "to": to_addrs, "count": len(paths)}
    except Exception as e:  # noqa: BLE001 — a failed send must not crash the caller
        print(f"FAILED to send report email: {e}")
        return {"status": "failed", "error": str(e)}
