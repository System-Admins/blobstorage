# Blob Browser

A static single-page web app for managing files in Azure Blob Storage. Runs entirely in the browser as a Storage Account **Static Website** — no back-end required. Authenticates via **OAuth 2.0 PKCE** against **Microsoft Entra ID** using only native browser APIs (no libraries).

## Features

- **Browse & navigate** — virtual folders, breadcrumbs, collapsible folder tree sidebar
- **Upload** — files and folders with drag-and-drop, progress bars, and overwrite conflict handling
- **Download** — single files or entire folders as ZIP
- **Create / Rename / Delete / Copy / Move** — folders and files (Contributors only)
- **View & Edit** — preview and edit text-based files in-browser (Contributors only)
- **Search** — current folder (instant filter) or everything (deep scan)
- **SAS generator** — User Delegation SAS URLs with configurable expiry, IP, and permissions
- **Storage picker** — switch between subscriptions, accounts, and containers via the ARM API
- **Email** — send file/folder links via Microsoft Graph (Mail.Send)
- **Audit metadata** — tracks uploader and last editor per blob
- **Auto RBAC detection** — UI adapts based on Reader vs. Contributor role; no config needed
- **Responsive** — desktop and mobile

---

## Prerequisites

- An **Azure subscription** with a **Storage Account** and at least one blob container
- Permission to create an **App Registration** in Microsoft Entra ID

---

## Setup

### Step 1 — Create an App Registration

1. [Azure Portal](https://portal.azure.com) → **Microsoft Entra ID** → **App registrations** → **New registration**
2. Name it (e.g. `Storage Explorer`), set to **Single tenant**
3. **Redirect URI** → platform **Single-page application (SPA)**:
   - `https://<storage-account>.<zone>.web.core.windows.net` (the URL for the static website in the Azure Storage Account, created in step 5)
4. Click **Register** and copy the **Application (client) ID** and **Directory (tenant) ID**

### Step 2 — Add API permissions

1. App Registration → **API permissions** → **Add a permission**
2. **Azure Storage** → Delegated → `user_impersonation` → Add
3. **Azure Service Management** → Delegated → `user_impersonation` → Add *(required for the Storage Picker; skip if using fixed account/container)*
4. Click **Grant admin consent** if required by your tenant

### Step 3 — Assign Storage RBAC roles

On the **Storage Account** → **Access control (IAM)** → **Add role assignment**:

| Role | Access |
|---|---|
| **Storage Blob Data Reader** | Browse, download, search, view, copy URL |
| **Storage Blob Data Contributor** | All of the above + upload, rename, delete, edit, create, SAS generation |

The app auto-detects the role and shows a **✏️ Writer** or **📖 Reader** badge.

### Step 4 — Configure CORS

On each **target** storage account you want to browse (not the one hosting the app):

**Storage Account** → **Resource sharing (CORS)** → **Blob service** → add a rule:

| Field | Value |
|---|---|
| Allowed origins | `https://<host-account>.<zone>.web.core.windows.net` |
| Allowed methods | `GET, HEAD, PUT, DELETE, POST, OPTIONS` |
| Allowed headers | `*` |
| Exposed headers | `*` |
| Max age | `86400` |

For local dev, add a second rule with origin `http://localhost:3000`.

### Step 5 — Enable Static Website hosting

1. **Storage Account** → **Static website** → **Enabled**
2. Index document: `index.html` — Error document: `index.html`
3. Note the **Primary endpoint** URL

### Step 6 — Update `config.js`

```js
const CONFIG = {
  auth: {
    clientId:  "YOUR-CLIENT-ID",
    tenantId:  "YOUR-TENANT-ID",
    // redirectUri is auto-detected from window.location.origin
  },
  storage: {
    accountName:   "",  // Leave empty to always show the storage picker
    containerName: "",  // Leave empty to always show the storage picker
  },
  app: {
    title:              "Blob Browser",
    allowDownload:      true,
    allowRename:        true,
    allowDelete:        true,
    allowStoragePicker: true,
    allowSas:           true,
    sasShowPermissions: false,  // false = always generate read-only SAS
    allowEmail:         true,
  },
  upload: {
    maxFileSizeMB: 0,  // 0 = no limit
  },
};
```

### Step 7 — Deploy

Upload the `src/` contents to the `$web` blob container:

```powershell
az storage blob upload-batch `
  --account-name <storage-account> `
  --destination '$web' `
  --source ./src `
  --overwrite
```

Open the **Primary endpoint** URL in your browser.

---


## Project structure

| File | Purpose |
|---|---|
| `config.js` | Settings — client/tenant IDs, storage target, feature flags |
| `auth.js` | OAuth 2.0 PKCE — sign-in, token exchange, silent refresh, Graph token |
| `arm.js` | ARM API — subscription/account/container discovery |
| `storage.js` | Blob Storage REST API — list, upload, download, rename, delete, SAS |
| `app.js` | UI — rendering, navigation, modals, upload queue, folder tree, search |
| `index.html` | HTML shell and all modal markup |
| `style.css` | Responsive stylesheet |

---

## How it works

**Authentication** — OAuth 2.0 Authorization Code with PKCE, implemented with `fetch` + `crypto.subtle`. Tokens are stored in `sessionStorage`. No MSAL or third-party libraries.

**Permission detection** — On login the app probes write access by attempting a PUT + DELETE of a zero-byte blob. Success → Contributor UI. Failure → Reader UI. Enforcement is always server-side.

**SAS generation** — Fetches a User Delegation Key via `POST`, then signs in-browser with HMAC-SHA256 (`crypto.subtle`). Max expiry: 7 days.

**Rename** — Uses Copy Blob + Delete Blob (no native rename in Azure). Folder renames are O(n) — one copy+delete per blob.

---

## Troubleshooting

| Symptom | Fix |
|---|---|
| `AADSTS50011` reply URL mismatch | Add the exact URL as an **SPA** Redirect URI in the App Registration |
| `AADSTS700054` response_type error | Delete the Web platform and re-add as **Single-page application** |
| 403 on listing/download | Assign **Storage Blob Data Reader** on the storage account |
| 403 on upload/rename/delete | Assign **Storage Blob Data Contributor** on the storage account |
| Upload button never appears | Confirm Contributor role and CORS allows `PUT` + `DELETE` |
| SAS 403 / CORS error | Add `POST` to CORS allowed methods on the target account |
| Storage picker shows nothing | Add **Azure Service Management** API permission + grant consent; assign **Reader** on the subscription |
| CORS errors in console | Add/fix the CORS rule on the **target** storage account (not the hosting account) |
| Blank page after deploy | Set index document to `index.html` in Static Website settings |
