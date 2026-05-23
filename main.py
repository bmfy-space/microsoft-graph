"""Email API server — single endpoint to fetch emails via Microsoft Graph API."""

from __future__ import annotations

import email as email_parser
import imaplib
from email.header import decode_header

import requests
from flask import Flask, jsonify, request, send_from_directory
from flask_cors import CORS

app = Flask(__name__)
CORS(app)

TOKEN_URL = "https://login.microsoftonline.com/common/oauth2/v2.0/token"
GRAPH_BASE = "https://graph.microsoft.com/v1.0"
PAGE_SIZE = 20
FOLDERS = ["inbox", "junkemail"]
GRAPH_REQUEST_ERROR = "Graph API request failed. Please retry."
TOKEN_REFRESH_ERROR = "Token refresh failed"
IMAP_HOST = "outlook.office365.com"
IMAP_FOLDER_MAP = {"inbox": "INBOX", "junkemail": "Junk Email"}

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


def token_error_message(data: dict) -> str:
    error = str(data.get("error") or "").strip()
    description = str(data.get("error_description") or "").strip()
    if error and description:
        return f"{error}: {description}"
    return error or description or TOKEN_REFRESH_ERROR


def _post_token(form: dict) -> tuple[dict | None, str | None]:
    resp = requests.post(
        TOKEN_URL,
        data=form,
        headers={"Content-Type": "application/x-www-form-urlencoded"},
        timeout=30,
    )
    try:
        data = resp.json()
    except ValueError:
        return None, TOKEN_REFRESH_ERROR
    if resp.ok and "access_token" in data:
        return data, None
    return None, token_error_message(data)


def get_access_token(client_id: str, refresh_token: str) -> tuple[str | None, str | None, str | None, str | None]:
    """Return (graph_token, opaque_token_for_imap, new_refresh_token, error)."""
    base_form = {"client_id": client_id, "grant_type": "refresh_token", "refresh_token": refresh_token}

    # Try with Mail.Read scope directly
    form = {**base_form, "scope": "Mail.Read"}
    data, err = _post_token(form)
    if data:
        return data["access_token"], None, data.get("refresh_token"), None

    # Scope failed — refresh without scope to get opaque token + new refresh_token
    data, err2 = _post_token(base_form)
    opaque_token = data.get("access_token") if data else None
    new_rt = data.get("refresh_token") if data else None
    working_rt = new_rt or refresh_token

    # Retry with scope using the new refresh_token
    form = {"client_id": client_id, "grant_type": "refresh_token", "refresh_token": working_rt, "scope": "Mail.Read"}
    data, err3 = _post_token(form)
    if data:
        return data["access_token"], None, new_rt or data.get("refresh_token"), None

    # Graph API failed — return opaque token for IMAP fallback
    return None, opaque_token, new_rt, err3 or err


def fetch_emails_page(
    access_token: str,
    folder: str,
    page: int = 1,
    page_size: int = PAGE_SIZE,
) -> tuple[list[dict], str | None]:
    url = f"{GRAPH_BASE}/me/mailFolders/{folder}/messages"
    params = {
        "$orderby": "receivedDateTime DESC",
        "$select": EMAIL_FIELDS,
        "$top": page_size,
        "$skip": (page - 1) * page_size,
    }
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Accept": "application/json",
    }
    resp = requests.get(url, params=params, headers=headers, timeout=30)
    if not resp.ok:
        raise requests.HTTPError(
            resp.json().get("error", {}).get("message", GRAPH_REQUEST_ERROR),
            response=resp,
        )
    data = resp.json()
    return data.get("value", []), data.get("@odata.nextLink")


def imap_connect(email: str, access_token: str) -> imaplib.IMAP4_SSL:
    auth_string = f"user={email}\x01auth=Bearer {access_token}\x01\x01"
    mail = imaplib.IMAP4_SSL(IMAP_HOST)
    mail.authenticate("XOAUTH2", lambda x: auth_string.encode())
    return mail


def _decode_header_value(raw: str | None) -> str:
    if not raw:
        return ""
    parts = decode_header(raw)
    decoded = []
    for data, charset in parts:
        if isinstance(data, bytes):
            decoded.append(data.decode(charset or "utf-8", errors="replace"))
        else:
            decoded.append(data)
    return "".join(decoded).strip()


def _extract_body(msg) -> dict:
    body_text = ""
    body_html = ""
    if msg.is_multipart():
        for part in msg.walk():
            ct = part.get_content_type()
            payload = part.get_payload(decode=True)
            if payload is None:
                continue
            charset = part.get_content_charset() or "utf-8"
            text = payload.decode(charset, errors="replace")
            if ct == "text/plain":
                body_text = body_text or text
            elif ct == "text/html":
                body_html = body_html or text
    else:
        payload = msg.get_payload(decode=True)
        if payload:
            charset = msg.get_content_charset() or "utf-8"
            text = payload.decode(charset, errors="replace")
            if msg.get_content_type() == "text/html":
                body_html = text
            else:
                body_text = text
    if body_html:
        return {"contentType": "html", "content": body_html}
    return {"contentType": "text", "content": body_text}


def parse_imap_message(msg_id: bytes, raw: bytes) -> dict:
    msg = email_parser.message_from_bytes(raw)
    from_hdr = msg.get("From", "")
    name, addr = email_parser.utils.parseaddr(from_hdr)
    if not name:
        name = _decode_header_value(from_hdr)

    to_list = []
    for val in (msg.get_all("To") or []):
        for part in val.split(","):
            n, a = email_parser.utils.parseaddr(part.strip())
            if n and a:
                to_list.append(f"{n} <{a}>")
            else:
                to_list.append(n or a)

    cc_list = []
    for val in (msg.get_all("Cc") or []):
        for part in val.split(","):
            n, a = email_parser.utils.parseaddr(part.strip())
            if n and a:
                cc_list.append(f"{n} <{a}>")
            else:
                cc_list.append(n or a)

    body = _extract_body(msg)
    preview = body["content"][:256].replace("\r\n", " ").replace("\n", " ") if body["content"] else ""

    return {
        "id": msg_id.decode(),
        "internetMessageId": msg.get("Message-ID", ""),
        "subject": _decode_header_value(msg.get("Subject")) or "(No subject)",
        "from": {"name": name, "address": addr},
        "to": to_list,
        "cc": cc_list,
        "date": msg.get("Date", ""),
        "preview": preview,
        "body": body,
    }


def fetch_emails_imap(
    mail: imaplib.IMAP4_SSL,
    folder: str,
    page: int = 1,
    page_size: int = PAGE_SIZE,
) -> list[dict]:
    imap_folder = IMAP_FOLDER_MAP.get(folder, folder)
    mail.select(imap_folder, readonly=True)

    status, data = mail.search(None, "ALL")
    if status != "OK" or not data[0]:
        return []

    msg_ids = data[0].split()
    msg_ids.reverse()  # newest first

    start = (page - 1) * page_size
    page_ids = msg_ids[start : start + page_size]

    results = []
    for mid in page_ids:
        status, msg_data = mail.fetch(mid, "(RFC822)")
        if status != "OK":
            continue
        raw = msg_data[0][1]
        results.append(parse_imap_message(mid, raw))
    return results


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
    folder = data.get("folder", "inbox")

    try:
        creds = parse_credentials(credentials, separator)
    except ValueError:
        return jsonify({"code": 400, "msg": "Invalid credentials format"}), 400

    try:
        graph_token, opaque_token, new_rt, token_error = get_access_token(creds["client_id"], creds["refresh_token"])
    except requests.RequestException:
        return jsonify({"code": 502, "msg": GRAPH_REQUEST_ERROR}), 502

    method = "graph"
    emails = None

    # 1. Graph API with scoped token
    if graph_token:
        try:
            items, _ = fetch_emails_page(graph_token, folder, page, page_size)
            emails = [build_email_json(item) for item in items]
        except requests.RequestException:
            pass

    # 2. Graph API with opaque token (works for some accounts without Mail.Read scope)
    if emails is None and opaque_token:
        try:
            items, _ = fetch_emails_page(opaque_token, folder, page, page_size)
            emails = [build_email_json(item) for item in items]
        except requests.RequestException:
            pass

    # 3. IMAP fallback
    if emails is None and opaque_token:
        method = "imap"
        try:
            mail = imap_connect(creds["email"], opaque_token)
            try:
                emails = fetch_emails_imap(mail, folder, page, page_size)
            finally:
                mail.logout()
        except Exception:
            pass

    if emails is None:
        return jsonify({"code": 401, "msg": token_error or TOKEN_REFRESH_ERROR}), 401

    resp_data = {
        "account": creds["email"],
        "folder": folder,
        "page": page,
        "pageSize": page_size,
        "total": len(emails),
        "method": method,
        "emails": emails,
    }
    if new_rt:
        resp_data["new_refresh_token"] = new_rt

    return jsonify({"code": 200, "msg": "success", "data": resp_data})


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

        try:
            graph_token, opaque_token, new_rt, token_error = get_access_token(creds["client_id"], creds["refresh_token"])
        except requests.RequestException:
            results.append({"email": creds["email"], "valid": False, "error": GRAPH_REQUEST_ERROR})
            continue

        result = {"email": creds["email"]}
        if graph_token:
            result["valid"] = True
            result["method"] = "graph"
        elif opaque_token:
            # Try Graph API with opaque token
            try:
                headers = {"Authorization": f"Bearer {opaque_token}", "Accept": "application/json"}
                resp = requests.get(f"{GRAPH_BASE}/me/messages?$top=1", headers=headers, timeout=15)
                if resp.ok:
                    result["valid"] = True
                    result["method"] = "graph"
                else:
                    raise Exception("not ok")
            except Exception:
                # Fall back to IMAP
                try:
                    mail = imap_connect(creds["email"], opaque_token)
                    mail.logout()
                    result["valid"] = True
                    result["method"] = "imap"
                except Exception:
                    result["valid"] = False
                    result["error"] = "Graph and IMAP both failed"
        else:
            result["valid"] = False
            result["error"] = token_error or TOKEN_REFRESH_ERROR
        if new_rt:
            result["new_refresh_token"] = new_rt
        results.append(result)

    return jsonify({"code": 200, "msg": "success", "data": results})


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=3000, debug=True)
