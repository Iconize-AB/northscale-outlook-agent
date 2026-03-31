# Outlook Calendar MCP — Setup Guide

This guide walks you through the three steps to get your ElevenLabs agent booking Outlook meetings:

1. Register an Azure App (get API credentials)
2. Deploy the MCP server to Vercel
3. Connect it to ElevenLabs

---

## Step 1 — Azure App Registration

You need to give the server permission to access your Outlook calendar via Microsoft Graph API.

### 1.1 Create the App Registration

1. Go to [https://portal.azure.com](https://portal.azure.com) and sign in with your Microsoft 365 / Outlook account.
2. Search for **"App registrations"** in the top search bar and open it.
3. Click **"+ New registration"**.
4. Fill in:
   - **Name:** `ElevenLabs Calendar MCP` (or any name)
   - **Supported account types:** *Accounts in this organizational directory only (Single tenant)*
   - **Redirect URI:** leave blank
5. Click **Register**.

### 1.2 Copy your IDs

After registration, you'll see an overview page. **Copy these two values** — you'll need them later:

- **Application (client) ID** → this is your `AZURE_CLIENT_ID`
- **Directory (tenant) ID** → this is your `AZURE_TENANT_ID`

### 1.3 Create a Client Secret

1. In the left menu, click **"Certificates & secrets"**.
2. Click **"+ New client secret"**.
3. Set a description (e.g. `vercel-mcp`) and an expiry (24 months is fine).
4. Click **Add**.
5. **Immediately copy the secret Value** (it won't be shown again) → this is your `AZURE_CLIENT_SECRET`.

### 1.4 Grant Calendar Permissions

1. In the left menu, click **"API permissions"**.
2. Click **"+ Add a permission"** → **Microsoft Graph** → **Application permissions**.
3. Search for and add these permissions:
   - `Calendars.ReadWrite`
   - `User.Read.All` (needed to look up the mailbox)
4. Click **"Grant admin consent for [your org]"** and confirm.
   - ⚠️ You need to be a Global Admin or have an admin do this step.
   - If you're using a personal Microsoft account, you can skip the admin consent step — the permissions auto-approve.

---

## Step 2 — Deploy to Vercel

### 2.1 Push to GitHub

1. Create a new GitHub repository (e.g. `outlook-calendar-mcp`).
2. Upload the project folder contents to the repo (or use `git push`).

### 2.2 Deploy on Vercel

1. Go to [https://vercel.com](https://vercel.com) and sign in (free account is fine).
2. Click **"Add New… → Project"**.
3. Import your GitHub repository.
4. Vercel will auto-detect it as a Node.js project. Click **Deploy**.

### 2.3 Add Environment Variables

After deploying, go to your project in Vercel → **Settings → Environment Variables** and add:

| Variable | Value |
|---|---|
| `AZURE_TENANT_ID` | The tenant ID from Step 1.2 |
| `AZURE_CLIENT_ID` | The client ID from Step 1.2 |
| `AZURE_CLIENT_SECRET` | The secret from Step 1.3 |
| `OUTLOOK_USER_EMAIL` | Your Outlook email address (e.g. `you@company.com`) |
| `MCP_SECRET` | A random secret string you choose (e.g. `my-secret-token-abc123`) |

After adding the variables, go to **Deployments → Redeploy** to apply them.

### 2.4 Note your MCP URL

Your MCP endpoint will be:
```
https://your-project-name.vercel.app/mcp
```

---

## Step 3 — Connect to ElevenLabs

1. In ElevenLabs, go to your Agent → **Integrations** → click **"Lägg till integration"** (Add integration).
2. Select **"Anpassad MCP-server"** (Custom MCP server).
3. Enter:
   - **URL:** `https://your-project-name.vercel.app/mcp`
   - **Authentication:** Bearer token
   - **Token:** the `MCP_SECRET` value you set in Vercel
4. Save the integration.

Your agent now has three tools available:

| Tool | What it does |
|---|---|
| `check_availability` | Finds free time slots in your calendar within a date range |
| `create_meeting` | Books a meeting, optionally with Teams link and invitees |
| `cancel_meeting` | Cancels a meeting by its event ID |

---

## Telling your agent how to use the tools

Add instructions like this to your ElevenLabs agent's system prompt:

```
You have access to the user's Outlook calendar. When asked about scheduling:
- Use check_availability to find open slots before suggesting times.
- Use create_meeting to book confirmed appointments. Always ask for confirmation before booking.
- Use cancel_meeting when asked to cancel. Always confirm the meeting details before cancelling.
- All times are in UTC — convert to the user's local timezone when communicating.
```

---

## Troubleshooting

**401 Unauthorized** — Check that the `MCP_SECRET` in Vercel matches what you entered in ElevenLabs.

**Graph API 403 Forbidden** — The admin consent hasn't been granted yet. Ask your Microsoft 365 admin to approve the permissions in Azure.

**Graph API 404** — Double-check that `OUTLOOK_USER_EMAIL` matches the exact email address of the calendar you want to access.
