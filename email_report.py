"""email_report.py — emails generated report PDFs via SMTP.

This is a single self-contained function. It reads each PDF from disk and
hands the bytes to Python's own email/smtplib standard-library modules, which
do MIME/base64 encoding internally, in memory, on whatever machine runs this
script. That encoding never has to pass through an LLM's context to happen —
it's the same plumbing every email client uses, just invisible.

No third-party package, no API key, no paid service — smtplib and email are
both in the Python standard library. All that's needed is a Gmail "App
Password" (Google Account -> Security -> 2-Step Verification -> App
Passwords — a 2-minute setup, and NOT your real Gmail password).

Credentials come from SMTP_USERNAME / SMTP_PASSWORD environment variables —
never committed, never printed. Sending is a deliberate no-op, not an error,
when they aren't configured, so mock/test-version runs and local development
never need email access at all.
"""
import os
import smtplib
import socket
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


def _connect_smtp_ipv4(host: str, port: int, timeout: int) -> smtplib.SMTP:
    """Connect to an SMTP host forcing IPv4, then hand back a normal SMTP
    object as if smtplib.SMTP(host, port) had connected directly.

    Some sandboxed environments have no IPv6 support in their network
    namespace at all. smtp.gmail.com resolves to both IPv4 and IPv6
    addresses, and Python's default connection logic can try IPv6 first --
    in such a sandbox that fails immediately with
    "OSError: [Errno 97] Address family not supported by protocol" (EAFNOSUPPORT),
    never falling back to the IPv4 address that would have worked fine.
    Resolving to a concrete IPv4 address ourselves and connecting to that
    sidesteps the family-selection question entirely.

    The original hostname is restored onto the connected object afterward,
    because starttls() needs it for TLS server-name verification against
    Gmail's certificate -- verifying against the raw IP would fail that
    check even though the underlying socket is IPv4.
    """
    ipv4_addr = socket.getaddrinfo(host, port, socket.AF_INET, socket.SOCK_STREAM)[0][4][0]
    server = smtplib.SMTP(timeout=timeout)
    server.connect(ipv4_addr, port)
    server._host = host  # noqa: SLF001 — restores the hostname for starttls()'s certificate check
    return server


def send_report_email(pdf_paths: list, to_addrs: list, subject: str = None, body: str = None) -> dict:
    """Email every PDF in pdf_paths as attachments, in one message, to to_addrs.

    Returns {"status": ..., ...}:
      "sent"    — delivered; includes "to" and "count"
      "skipped" — not an error; includes "reason" (no credentials/recipients/PDFs)
      "failed"  — includes "error"
    """
    if not pdf_paths:
        return {"status": "skipped", "reason": "no PDFs to send"}

    if not to_addrs:
        print("no recipients configured — skipping email delivery")
        return {"status": "skipped", "reason": "no recipients"}

    # .strip() because a password pasted into .env or a secrets UI routinely
    # picks up a trailing newline, which SMTP auth then rejects.
    username = (os.environ.get("SMTP_USERNAME") or "").strip()
    password = (os.environ.get("SMTP_PASSWORD") or "").strip()
    if not username or not password:
        print("SMTP_USERNAME/SMTP_PASSWORD not set — skipping email delivery")
        return {"status": "skipped", "reason": "no SMTP credentials configured"}

    host = os.environ.get("SMTP_HOST", "smtp.gmail.com")
    port = int(os.environ.get("SMTP_PORT", "587"))

    try:
        msg = MIMEMultipart()
        msg["From"] = username
        msg["To"] = ", ".join(to_addrs)
        msg["Subject"] = subject or "Quiet List Exchange Activity Report"
        msg.attach(MIMEText(
            body or "Attached: the latest Quiet List Exchange Activity Report(s).",
            "plain",
        ))

        # A missing/unreadable PDF (e.g. from a bad path) must fail cleanly
        # too, not just an SMTP-connection problem — this whole block is one
        # try so nothing here can crash the caller.
        for path in pdf_paths:
            with open(path, "rb") as f:
                part = MIMEApplication(f.read(), _subtype="pdf")
            part.add_header("Content-Disposition", "attachment", filename=os.path.basename(path))
            msg.attach(part)

        server = _connect_smtp_ipv4(host, port, timeout=30)
        try:
            server.starttls()
            server.login(username, password)
            server.sendmail(username, to_addrs, msg.as_string())
        finally:
            server.quit()
        print(f"emailed {len(pdf_paths)} report(s) to {', '.join(to_addrs)}")
        return {"status": "sent", "to": to_addrs, "count": len(pdf_paths)}
    except Exception as e:  # noqa: BLE001 — a failed send must not crash the caller
        print(f"FAILED to send report email: {e}")
        return {"status": "failed", "error": str(e)}
