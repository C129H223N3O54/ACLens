# ACLens — Manual SharePoint Setup Guide

This guide walks you through manually creating an Azure AD App Registration for ACLens. Use this if the automatic Device Code setup did not work.

---

## Prerequisites

- You need **Application Administrator** or **Global Administrator** role in your Microsoft 365 tenant
- Access to the [Azure Portal](https://portal.azure.com)

---

## Step 1 — Create the App Registration

1. Go to [portal.azure.com](https://portal.azure.com)
2. Search for **"App registrations"** in the top search bar
3. Click **+ New registration**
4. Fill in:
   - **Name:** `ACLens`
   - **Supported account types:** `Accounts in this organizational directory only`
   - **Redirect URI:** leave empty
5. Click **Register**

> After registration you will see the **Application (client) ID** and **Directory (tenant) ID** — copy both, you will need them in ACLens.

---

## Step 2 — Add API Permissions

1. In your new app registration, click **API permissions** in the left menu
2. Click **+ Add a permission**
3. Select **Microsoft Graph**
4. Select **Application permissions** (not Delegated)
5. Search for and add the following permissions:

| Permission | Description |
|---|---|
| `Sites.Read.All` | Read items in all site collections |
| `Sites.FullControl.All` | Required to read site-level permission assignments |
| `User.Read.All` | Read all user profiles (for resolving names) |
| `GroupMember.Read.All` | Read group memberships |

6. Click **Add permissions**
7. Click **Grant admin consent for [Your Tenant]** — confirm with **Yes**

> The permissions will show a green checkmark once consent is granted.

---

## Step 3 — Create a Client Secret

1. Click **Certificates & secrets** in the left menu
2. Click **+ New client secret**
3. Add a description (e.g. `ACLens Secret`)
4. Set expiry — **24 months** recommended
5. Click **Add**
6. **Copy the secret Value immediately** — it will only be shown once

---

## Step 4 — Enter Credentials in ACLens

1. Open ACLens
2. Click the **SharePoint** tab
3. Click **Setup / Reconnect**
4. Click **Manual Setup** tab
5. Enter:
   - **Tenant ID** — from Step 1 (Directory ID)
   - **Client ID** — from Step 1 (Application ID)
   - **Client Secret** — from Step 3
6. Click **Save & Connect**

ACLens will store the credentials encrypted on your local machine at `sp_credentials.xml` next to the script.

---

## Verifying the Connection

After saving, the SharePoint tab should show:

```
✔  Connected
Tenant: xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx  |  Client: xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
```

---

## Required Permissions Summary

| API | Permission | Type | Purpose |
|---|---|---|---|
| Microsoft Graph | `Sites.Read.All` | Application | Read SharePoint site contents |
| Microsoft Graph | `Sites.FullControl.All` | Application | Read permission assignments on sites |
| Microsoft Graph | `User.Read.All` | Application | Resolve user display names |
| Microsoft Graph | `GroupMember.Read.All` | Application | Read group members |

---

## Troubleshooting

| Problem | Solution |
|---|---|
| "Insufficient privileges" error | Make sure admin consent was granted for all permissions (green checkmarks in API permissions) |
| "Application not found" | Check that the Client ID is correct and matches the App Registration |
| "Invalid client secret" | The secret may have expired or been entered incorrectly — create a new one |
| Site not found | Make sure the Site URL is exactly as shown in SharePoint (copy from browser address bar) |
| 403 Forbidden on folders | `Sites.FullControl.All` may not have been consented — re-grant admin consent |

---

## Security Notes

- The client secret is stored **encrypted** using Windows Data Protection API (DPAPI) — it is tied to your Windows user account and cannot be decrypted on another machine
- ACLens only requests **read** permissions — it cannot modify permissions or content
- The `sp_credentials.xml` file should not be shared or committed to version control — it is listed in `.gitignore`

---

*Back to [README](../README.md)*
