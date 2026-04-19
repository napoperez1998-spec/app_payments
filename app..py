"""
Payment Mailer (Outlook Edition)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Streamlit app that reads payment-related events from your Outlook
calendar and sends reminder emails — all through Microsoft Graph API.

Setup:
  1. Register an app at https://portal.azure.com → Azure Active Directory
     → App registrations → New registration.
  2. Set redirect URI to  http://localhost:8501  (type: Web).
  3. Under API permissions add:
       • Calendars.Read
       • Mail.Send
     and grant admin consent (or user consent if allowed).
  4. Copy Application (client) ID and Tenant ID into the sidebar
     (or set env vars OUTLOOK_CLIENT_ID / OUTLOOK_TENANT_ID).
  5. pip install -r requirements.txt
  6. streamlit run app.py
"""

import os
import json
import datetime as dt
from pathlib import Path

import msal
import requests
import streamlit as st
from dateutil import parser as dtparser

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
GRAPH_BASE = "https://graph.microsoft.com/v1.0"
SCOPES = ["Calendars.Read", "Mail.Send"]
AUTHORITY_TPL = "https://login.microsoftonline.com/{tenant}"
REDIRECT_URI = "http://localhost:8501"
TOKEN_CACHE_PATH = Path("msal_token_cache.json")

PAYMENT_KEYWORDS = [
    "payment", "invoice", "billing", "due", "fee", "subscription",
    "pago", "factura", "cobro", "cuota", "vencimiento", "suscripción",
]

# ---------------------------------------------------------------------------
# Auth helpers
# ---------------------------------------------------------------------------

def _load_cache() -> msal.SerializableTokenCache:
    cache = msal.SerializableTokenCache()
    if TOKEN_CACHE_PATH.exists():
        cache.deserialize(TOKEN_CACHE_PATH.read_text())
    return cache


def _save_cache(cache: msal.SerializableTokenCache):
    if cache.has_state_changed:
        TOKEN_CACHE_PATH.write_text(cache.serialize())


def _build_app(client_id: str, tenant_id: str, cache: msal.SerializableTokenCache):
    return msal.PublicClientApplication(
        client_id,
        authority=AUTHORITY_TPL.format(tenant=tenant_id),
        token_cache=cache,
    )


def get_access_token(client_id: str, tenant_id: str) -> str | None:
    """Try to get a token silently; return None if interactive login needed."""
    cache = _load_cache()
    app = _build_app(client_id, tenant_id, cache)
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])
        if result and "access_token" in result:
            _save_cache(cache)
            return result["access_token"]
    return None


def interactive_login(client_id: str, tenant_id: str) -> str:
    """Run device-code flow (works in any environment, no browser redirect needed)."""
    cache = _load_cache()
    app = _build_app(client_id, tenant_id, cache)
    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        raise RuntimeError(f"Device flow failed: {json.dumps(flow, indent=2)}")

    st.info(
        f"👉 Go to **[{flow['verification_uri']}]({flow['verification_uri']})** "
        f"and enter code **`{flow['user_code']}`**"
    )
    st.caption("Waiting for you to sign in …")

    result = app.acquire_token_by_device_flow(flow)
    if "access_token" not in result:
        raise RuntimeError(f"Authentication failed: {result.get('error_description', result)}")

    _save_cache(cache)
    return result["access_token"]


def graph_headers(token: str) -> dict:
    return {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}


# ---------------------------------------------------------------------------
# Calendar helpers
# ---------------------------------------------------------------------------

def fetch_payment_events(token: str, days_ahead: int = 30,
                         custom_keywords: list[str] | None = None) -> list[dict]:
    """Return Outlook calendar events matching payment keywords."""
    now = dt.datetime.utcnow()
    time_min = now.strftime("%Y-%m-%dT%H:%M:%SZ")
    time_max = (now + dt.timedelta(days=days_ahead)).strftime("%Y-%m-%dT%H:%M:%SZ")

    url = (
        f"{GRAPH_BASE}/me/calendarView"
        f"?startDateTime={time_min}&endDateTime={time_max}"
        f"&$top=100&$orderby=start/dateTime"
        f"&$select=subject,bodyPreview,start,end,webLink"
    )

    resp = requests.get(url, headers=graph_headers(token))
    resp.raise_for_status()
    all_events = resp.json().get("value", [])

    keywords = PAYMENT_KEYWORDS + (custom_keywords or [])
    keywords_lower = [k.lower() for k in keywords]

    matched = []
    for ev in all_events:
        text = (ev.get("subject", "") + " " + ev.get("bodyPreview", "")).lower()
        if any(kw in text for kw in keywords_lower):
            matched.append(ev)

    return matched


def parse_event_datetime(event: dict) -> str:
    raw = event.get("start", {}).get("dateTime", "")
    try:
        parsed = dtparser.parse(raw)
        return parsed.strftime("%A, %B %d %Y – %I:%M %p")
    except Exception:
        return raw


# ---------------------------------------------------------------------------
# Mail helpers
# ---------------------------------------------------------------------------

def send_email(token: str, to: str, subject: str, body_html: str,
               cc: str = "", bcc: str = ""):
    """Send an email via Microsoft Graph / Outlook."""
    to_recipients = [{"emailAddress": {"address": a.strip()}}
                     for a in to.split(",") if a.strip()]
    cc_recipients = [{"emailAddress": {"address": a.strip()}}
                     for a in cc.split(",") if a.strip()]
    bcc_recipients = [{"emailAddress": {"address": a.strip()}}
                      for a in bcc.split(",") if a.strip()]

    payload = {
        "message": {
            "subject": subject,
            "body": {"contentType": "HTML", "content": body_html},
            "toRecipients": to_recipients,
            "ccRecipients": cc_recipients,
            "bccRecipients": bcc_recipients,
        }
    }

    resp = requests.post(
        f"{GRAPH_BASE}/me/sendMail",
        headers=graph_headers(token),
        json=payload,
    )
    resp.raise_for_status()
    return True


# ---------------------------------------------------------------------------
# Email template
# ---------------------------------------------------------------------------

def build_payment_email_body(events: list[dict], custom_note: str = "") -> str:
    rows = ""
    for ev in events:
        date_str = parse_event_datetime(ev)
        summary = ev.get("subject", "(no title)")
        description = ev.get("bodyPreview", "—")
        rows += f"""
        <tr>
            <td style="padding:12px 16px; border-bottom:1px solid #ededed;
                        font-weight:600; color:#1b1b3a;">{summary}</td>
            <td style="padding:12px 16px; border-bottom:1px solid #ededed;
                        color:#555;">{date_str}</td>
            <td style="padding:12px 16px; border-bottom:1px solid #ededed;
                        color:#555; max-width:260px;">{description[:140]}</td>
        </tr>"""

    note_block = ""
    if custom_note.strip():
        note_block = f"""
        <tr><td colspan="3" style="padding:20px 16px 8px; color:#333;
                font-size:14px;">
            <strong>Note:</strong> {custom_note}
        </td></tr>"""

    return f"""
    <div style="font-family:'Segoe UI',Calibri,Roboto,Helvetica,Arial,sans-serif;
                max-width:700px; margin:auto; color:#222;">
        <div style="background:linear-gradient(135deg,#0f6cbd 0%,#2899f5 100%);
                    padding:28px 32px; border-radius:10px 10px 0 0;">
            <h1 style="margin:0; color:#fff; font-size:22px; letter-spacing:-.3px;">
                💳 Payment Reminder
            </h1>
            <p style="margin:6px 0 0; color:#cce4ff; font-size:14px;">
                Upcoming payments from your Outlook Calendar
            </p>
        </div>
        <table style="width:100%; border-collapse:collapse; background:#fff;
                      border:1px solid #e0e0e0;">
            <thead>
                <tr style="background:#f4f6f9;">
                    <th style="text-align:left; padding:12px 16px;
                               color:#0f6cbd; font-size:13px;">Event</th>
                    <th style="text-align:left; padding:12px 16px;
                               color:#0f6cbd; font-size:13px;">Date</th>
                    <th style="text-align:left; padding:12px 16px;
                               color:#0f6cbd; font-size:13px;">Details</th>
                </tr>
            </thead>
            <tbody>
                {rows}
                {note_block}
            </tbody>
        </table>
        <p style="font-size:12px; color:#999; text-align:center;
                  padding:18px; margin:0;">
            Sent via <em>Payment Mailer</em> · Outlook Edition
        </p>
    </div>"""


# ---------------------------------------------------------------------------
# Streamlit UI
# ---------------------------------------------------------------------------

def main():
    st.set_page_config(page_title="Payment Mailer · Outlook", page_icon="💳", layout="wide")

    # ── Custom CSS ────────────────────────────────────────────────────────
    st.markdown("""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;700&display=swap');
        html, body, [class*="st-"] { font-family: 'DM Sans', sans-serif; }
        .block-container { max-width: 960px; padding-top: 2rem; }
        h1, h2, h3 { color: #0f6cbd; }
        .stButton > button {
            background: linear-gradient(135deg, #0f6cbd, #2899f5);
            color: white; border: none; border-radius: 8px;
            padding: 0.55rem 1.6rem; font-weight: 600;
            transition: transform .15s, box-shadow .15s;
        }
        .stButton > button:hover {
            transform: translateY(-1px);
            box-shadow: 0 4px 14px rgba(15,108,189,.35);
        }
        div[data-testid="stExpander"] {
            border: 1px solid #e0e0e0; border-radius: 10px;
        }
    </style>""", unsafe_allow_html=True)

    # ── Header ────────────────────────────────────────────────────────────
    st.markdown("## 💳 Payment Mailer — Outlook Edition")
    st.caption("Scan your Outlook calendar for payment events and send email reminders.")

    # ── Sidebar: Azure app settings ───────────────────────────────────────
    with st.sidebar:
        st.header("🔧 Azure App Settings")
        client_id = st.text_input(
            "Application (client) ID",
            value=os.environ.get("OUTLOOK_CLIENT_ID", ""),
            type="password",
        )
        tenant_id = st.text_input(
            "Directory (tenant) ID",
            value=os.environ.get("OUTLOOK_TENANT_ID", "common"),
            help='Use "common" to allow any Microsoft account, or paste your tenant ID.',
        )

        st.divider()
        st.header("⚙️ Scan Settings")
        days_ahead = st.slider("Look-ahead window (days)", 1, 90, 30)
        extra_kw = st.text_input(
            "Extra keywords (comma-separated)",
            placeholder="rent, subscription, premium",
        )
        custom_keywords = [k.strip() for k in extra_kw.split(",") if k.strip()] if extra_kw else []

        st.divider()
        st.caption(
            "Payment Mailer scans your Outlook calendar for events "
            "matching payment keywords and sends styled reminder "
            "emails via Microsoft Graph."
        )

    if not client_id:
        st.warning("Enter your Azure **Application (client) ID** in the sidebar to get started.")
        st.stop()

    # ── Authentication ────────────────────────────────────────────────────
    token = get_access_token(client_id, tenant_id)

    if token:
        st.success("✅ Connected to Outlook", icon="🔗")
    else:
        st.info("Sign in with your Microsoft account to access Calendar & Mail.")
        if st.button("🔐 Sign in with Microsoft"):
            try:
                token = interactive_login(client_id, tenant_id)
                st.session_state["ms_token"] = token
                st.success("✅ Authenticated!")
                st.rerun()
            except Exception as exc:
                st.error(f"❌ Authentication failed: {exc}")
                st.stop()
        st.stop()

    # keep token handy across reruns
    if token:
        st.session_state["ms_token"] = token
    token = st.session_state.get("ms_token", token)

    # ── Fetch events ──────────────────────────────────────────────────────
    st.markdown("---")
    col_fetch, _ = st.columns([1, 3])
    with col_fetch:
        fetch_btn = st.button("🔄 Scan Calendar")

    if fetch_btn or "events" in st.session_state:
        if fetch_btn:
            with st.spinner("Scanning Outlook calendar…"):
                try:
                    events = fetch_payment_events(token, days_ahead, custom_keywords)
                except requests.HTTPError as exc:
                    st.error(f"❌ Graph API error: {exc.response.text}")
                    st.stop()
            st.session_state["events"] = events

        events: list[dict] = st.session_state.get("events", [])

        if not events:
            st.warning(
                f"No payment-related events found in the next {days_ahead} days. "
                "Try adding extra keywords in the sidebar."
            )
            st.stop()

        st.markdown(f"### 📅 Found {len(events)} payment event(s)")

        # ── Event cards ──────────────────────────────────────────────────
        selected_indices: list[int] = []
        for i, ev in enumerate(events):
            date_str = parse_event_datetime(ev)
            summary = ev.get("subject", "(no title)")
            desc = ev.get("bodyPreview", "—")

            with st.expander(f"**{summary}** — {date_str}"):
                st.write(desc if desc.strip() else "_No description._")
                link = ev.get("webLink")
                if link:
                    st.markdown(f"[Open in Outlook ↗]({link})")

            if st.checkbox("Include in email", value=True, key=f"ev_{i}"):
                selected_indices.append(i)

        # ── Compose email ─────────────────────────────────────────────────
        st.markdown("---")
        st.markdown("### ✉️ Compose Email")

        col1, col2 = st.columns(2)
        with col1:
            to_addr = st.text_input("To", placeholder="recipient@outlook.com")
            cc_addr = st.text_input("CC (optional)")
        with col2:
            bcc_addr = st.text_input("BCC (optional)")
            subject = st.text_input(
                "Subject",
                value=f"Payment Reminder – {len(selected_indices)} upcoming event(s)",
            )

        custom_note = st.text_area(
            "Add a personal note (optional)",
            placeholder="e.g. Please confirm receipt of this reminder.",
        )

        # ── Preview ───────────────────────────────────────────────────────
        selected_events = [events[i] for i in selected_indices]

        with st.expander("👁️ Preview email"):
            preview_html = build_payment_email_body(selected_events, custom_note)
            st.components.v1.html(preview_html, height=400, scrolling=True)

        # ── Send ──────────────────────────────────────────────────────────
        send_disabled = not to_addr or not selected_events
        if st.button("🚀 Send Email", disabled=send_disabled, type="primary"):
            body_html = build_payment_email_body(selected_events, custom_note)
            try:
                with st.spinner("Sending via Outlook…"):
                    send_email(token, to_addr, subject, body_html,
                               cc=cc_addr, bcc=bcc_addr)
                st.success("✅ Email sent successfully!")
                st.balloons()
            except requests.HTTPError as exc:
                st.error(f"❌ Failed to send: {exc.response.text}")
            except Exception as exc:
                st.error(f"❌ Failed to send: {exc}")


if __name__ == "__main__":
    main()
