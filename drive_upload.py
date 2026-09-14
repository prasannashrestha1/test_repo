"""drive_upload.py — uploads generated report files (PDF and .docx alike) to
Google Drive.

A cloud routine's filesystem does not persist between runs, so a report that
only lands in ./reports is lost the moment the session ends. This uploads it
somewhere durable instead.

Auth is a Google Cloud service account, NOT the Drive connector a Claude
session itself might have — a plain Python subprocess has no access to that
connector; it needs its own credential. The service account's key (JSON) is
supplied via the GOOGLE_SERVICE_ACCOUNT_JSON environment variable — the raw
JSON content, never a file path baked into config, and never committed.

Uploading is a deliberate no-op, not an error, when that variable isn't set —
so mock/test-version runs and local development never need Drive access.

One-time setup (see README.md for the full walkthrough):
  1. Create a Google Cloud service account, enable the Drive API for it.
  2. Generate a JSON key for it.
  3. Share the target Drive folder (drive_reports_folder_id in config.json)
     with the service account's email address (looks like
     xxx@project-id.iam.gserviceaccount.com), Editor access.
  4. Set GOOGLE_SERVICE_ACCOUNT_JSON to that key's raw JSON content.
"""
import json
import mimetypes
import os


def _get_drive_service():
    creds_json = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")
    if not creds_json:
        return None

    # Imported lazily so these packages are only required when this feature
    # is actually configured — a mock/test-version run never needs them.
    from google.oauth2 import service_account
    from googleapiclient.discovery import build

    info = json.loads(creds_json)
    credentials = service_account.Credentials.from_service_account_info(
        info, scopes=["https://www.googleapis.com/auth/drive.file"]
    )
    return build("drive", "v3", credentials=credentials, cache_discovery=False)


def upload_reports(file_paths: list, folder_id: str) -> list:
    """Upload each file (PDF, .docx, whatever) to the given Drive folder.

    Returns one {path, status, ...} dict per input path. status is
    "uploaded" (with drive_id/link), "skipped" (no credentials configured —
    not an error), or "failed" (with an "error" message).
    """
    if not file_paths:
        return []

    if not os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON"):
        print("GOOGLE_SERVICE_ACCOUNT_JSON not set — skipping Drive upload "
              "(files remain only in ./reports)")
        return [{"path": p, "status": "skipped"} for p in file_paths]

    if not folder_id:
        print("drive_reports_folder_id missing from config.json — skipping Drive upload")
        return [{"path": p, "status": "skipped"} for p in file_paths]

    try:
        service = _get_drive_service()
    except Exception as e:  # noqa: BLE001 — bad/malformed credentials must fail each file cleanly, not crash the whole scheduled run
        print(f"FAILED to initialize Drive credentials: {e}")
        return [{"path": p, "status": "failed", "error": f"credentials error: {e}"} for p in file_paths]

    from googleapiclient.http import MediaFileUpload

    results = []
    for path in file_paths:
        name = os.path.basename(path)
        try:
            mime_type, _ = mimetypes.guess_type(path)
            metadata = {"name": name, "parents": [folder_id]}
            media = MediaFileUpload(path, mimetype=mime_type or "application/octet-stream",
                                     resumable=True)
            uploaded = service.files().create(
                body=metadata, media_body=media, fields="id, webViewLink"
            ).execute()
            print(f"uploaded {name} -> Drive file {uploaded['id']}")
            results.append({
                "path": path, "status": "uploaded",
                "drive_id": uploaded["id"], "link": uploaded.get("webViewLink"),
            })
        except Exception as e:  # noqa: BLE001 — one bad upload must not stop the rest
            print(f"FAILED to upload {name}: {e}")
            results.append({"path": path, "status": "failed", "error": str(e)})
    return results
