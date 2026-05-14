"""Email API server — single endpoint to fetch emails via Microsoft Graph API."""

from __future__ import annotations

from datetime import datetime

import requests
from flask import Flask, jsonify, request, send_from_directory
from flask_cors import CORS

app = Flask(__name__)
CORS(app)

TOKEN_URL = "https://login.microsoftonline.com/common/oauth2/v2.0/token"
GRAPH_BASE = "https://graph.microsoft.com/v1.0"
SCOPE = "https://graph.microsoft.com/.default"
PAGE_SIZE = 20
FOLDERS = ["inbox", "junkemail"]

EMAIL_FIELDS = (
    "id,internetMessageId,subject,bodyPreview,receivedDateTime,"
    "from,toRecipients,ccRecipients,body"
)


def parse_credentials(raw: str, separator: str = "----") -> dict[str, str]:
    email, password, client_id, refresh_token = raw.split(separator, 3)
    return {
        "email": email,
        "password": password,
        "client_id": client_id,
        "refresh_token": refresh_token,
    }


def get_access_token(client_id: str, refresh_token: str) -> str | None:
    resp = requests.post(
        TOKEN_URL,
        data={
            "client_id": client_id,
            "grant_type": "refresh_token",
            "refresh_token": refresh_token,
            "scope": SCOPE,
        },
        headers={"Content-Type": "application/x-www-form-urlencoded"},
        timeout=30,
    )
    data = resp.json()
    if not resp.ok or "access_token" not in data:
        return None
    return data["access_token"]


def _fetch_folder_page(
    access_token: str, folder: str, page_size: int, url: str | None = None
) -> tuple[list[dict], str | None]:
    if url is None:
        url = f"{GRAPH_BASE}/me/mailFolders/{folder}/messages"
        params = {
            "$orderby": "receivedDateTime DESC",
            "$select": EMAIL_FIELDS,
            "$top": page_size,
        }
    else:
        params = None

    headers = {
        "Authorization": f"Bearer {access_token}",
        "Accept": "application/json",
    }
    resp = requests.get(url, params=params, headers=headers, timeout=30)
    resp.raise_for_status()
    data = resp.json()
    return data.get("value", []), data.get("@odata.nextLink")


def fetch_emails(
    access_token: str,
    page: int = 1,
    page_size: int = PAGE_SIZE,
    folders: list[str] | None = None,
) -> list[dict]:
    target_count = page * page_size + 1
    merged: dict[str, dict] = {}

    for folder in (folders or FOLDERS):
        next_link = None
        fetched = 0
        while True:
            items, next_link = _fetch_folder_page(
                access_token, folder, page_size, url=next_link
            )
            for item in items:
                uid = item.get("id")
                if uid and uid not in merged:
                    merged[uid] = item
            fetched += len(items)
            if not next_link or fetched >= target_count:
                break

    all_items = sorted(
        merged.values(),
        key=lambda x: x.get("receivedDateTime") or "",
        reverse=True,
    )

    start = (page - 1) * page_size
    return all_items[start : start + page_size]


def format_recipients(recipients: list[dict] | None) -> list[str]:
    if not recipients:
        return []
    result = []
    for r in recipients:
        addr = r.get("emailAddress", {})
        name = (addr.get("name") or "").strip()
        email = (addr.get("address") or "").strip()
        if name and email:
            result.append(f"{name} <{email}>")
        else:
            result.append(name or email)
    return result


def build_email_json(item: dict) -> dict:
    addr = item.get("from", {}).get("emailAddress", {}) if item.get("from") else {}
    body = item.get("body", {})
    return {
        "id": item.get("id"),
        "internetMessageId": item.get("internetMessageId"),
        "subject": item.get("subject") or "(No subject)",
        "from": {
            "name": (addr.get("name") or "").strip(),
            "address": (addr.get("address") or "").strip(),
        },
        "to": format_recipients(item.get("toRecipients")),
        "cc": format_recipients(item.get("ccRecipients")),
        "date": item.get("receivedDateTime"),
        "preview": item.get("bodyPreview") or "",
        "body": {
            "contentType": body.get("contentType", "text"),
            "content": body.get("content") or "",
        },
    }


@app.route("/")
def index():
    return send_from_directory("static", "index.html")


@app.route("/api/email", methods=["POST"])
def api_email():
    data = request.get_json(silent=True) or {}

    credentials = data.get("credentials", "")
    if not credentials:
        return jsonify({"code": 400, "msg": "credentials is required"}), 400

    separator = data.get("separator", "----")
    page = int(data.get("page", 1))
    page_size = int(data.get("pageSize", PAGE_SIZE))
    folders = data.get("folders")

    try:
        creds = parse_credentials(credentials, separator)
    except ValueError:
        return jsonify({"code": 400, "msg": "Invalid credentials format"}), 400

    access_token = get_access_token(creds["client_id"], creds["refresh_token"])
    if access_token is None:
        return jsonify({"code": 401, "msg": "Failed to refresh token"}), 401

    try:
        items = fetch_emails(access_token, page, page_size, folders)
    except requests.HTTPError as e:
        return jsonify({"code": 502, "msg": f"Graph API error: {e}"}), 502

    return jsonify({
        "code": 200,
        "msg": "success",
        "data": {
            "account": creds["email"],
            "page": page,
            "pageSize": page_size,
            "total": len(items),
            "emails": [build_email_json(item) for item in items],
        },
    })


@app.route("/api/check", methods=["POST"])
def api_check():
    data = request.get_json(silent=True) or {}

    credentials = data.get("credentials", [])
    if isinstance(credentials, str):
        credentials = [credentials]
    if not credentials:
        return jsonify({"code": 400, "msg": "credentials is required"}), 400

    separator = data.get("separator", "----")
    results = []

    for raw in credentials:
        raw = raw.strip()
        if not raw:
            continue
        try:
            creds = parse_credentials(raw, separator)
        except ValueError:
            results.append({"email": raw.split(separator)[0], "valid": False, "error": "Invalid format"})
            continue

        token = get_access_token(creds["client_id"], creds["refresh_token"])
        if token:
            results.append({"email": creds["email"], "valid": True})
        else:
            results.append({"email": creds["email"], "valid": False, "error": "Token refresh failed"})

    return jsonify({"code": 200, "msg": "success", "data": results})


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=3000, debug=True)
