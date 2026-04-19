# 💳 Payment Mailer — Outlook Edition

A Streamlit app that scans your **Outlook calendar** for payment-related events and sends styled HTML reminder emails via **Microsoft Graph API**.

---

## Features

- **Outlook Calendar scanning** — Reads events via Microsoft Graph's `/calendarView` endpoint.
- **Smart keyword matching** — Detects payment events in English & Spanish: *payment, invoice, billing, due, fee, pago, factura, cobro, cuota, vencimiento*, plus your own custom keywords.
- **Send via Outlook** — Emails sent through Microsoft Graph's `/me/sendMail` (your Outlook account).
- **Device-code auth** — Works everywhere, no browser-redirect hassle; just visit a URL and paste a code.
- **Styled HTML emails** — Professional template with event table, personal notes, CC/BCC.
- **Live preview** — See the email before sending.

---

## Setup

### 1. Register an Azure App

1. Go to [Azure Portal → App registrations](https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade).
2. Click **New registration**.
3. Name it (e.g. *Payment Mailer*), set **Supported account types** to *Accounts in any organizational directory and personal Microsoft accounts*.
4. Set **Redirect URI** → Web → `http://localhost:8501` (optional for device-code flow).
5. Click **Register**.

### 2. Add API Permissions

In your app registration:

1. Go to **API permissions → Add a permission → Microsoft Graph → Delegated permissions**.
2. Add:
   - `Calendars.Read`
   - `Mail.Send`
3. Click **Grant admin consent** (or have your admin do it).

### 3. Copy your IDs

From the **Overview** page of your app registration, copy:

| Field | Where to paste |
|-------|---------------|
| **Application (client) ID** | Sidebar → "Application (client) ID" |
| **Directory (tenant) ID** | Sidebar → "Directory (tenant) ID" (or use `common`) |

You can also set them as environment variables:
```bash
export OUTLOOK_CLIENT_ID="your-client-id"
export OUTLOOK_TENANT_ID="your-tenant-id"   # or "common"
```

### 4. Install & Run

```bash
cd payment_mailer_outlook
pip install -r requirements.txt
streamlit run app.py
```

---

## Usage

1. Enter your **Client ID** and **Tenant ID** in the sidebar.
2. Click **Sign in with Microsoft** — you'll get a device code to enter at `microsoft.com/devicelogin`.
3. Adjust the look-ahead window and add extra keywords in the sidebar.
4. Click **🔄 Scan Calendar** to find payment events.
5. Select events, fill in recipients, preview, and hit **🚀 Send Email**.

---

## File Structure

```
payment_mailer_outlook/
├── app.py                 # Main Streamlit application
├── requirements.txt       # Python dependencies
├── msal_token_cache.json  # Auto-generated after first sign-in
└── README.md
```
