# Azure Blob Storage Explorer

A static single-page web application for browsing, downloading, uploading, and managing files in Azure Blob Storage. Runs entirely in the browser as a Storage Account **Static Website** — no back-end or server required.

Authentication is handled natively using the **OAuth 2.0 Authorization Code Flow with PKCE** against **Microsoft Entra ID** (Azure AD). No third-party libraries are used.

---

## Features

| Feature | Details |
|---|---|
| 🔐 Authentication | Native OAuth 2.0 PKCE — no external libraries |
| 🔄 Silent token refresh | Refresh tokens kept in `sessionStorage`; re-login only when session fully expires |
| 📁 Folder navigation | Virtual-directory browsing with breadcrumb trail, Up button, and folder tree sidebar |
| 🌳 Folder tree sidebar | Collapsible tree panel showing the full container structure; clicking a node navigates directly |
| 📋 / ⊞ Views | Toggle between list view (table) and grid view |
| ⬇️ Download | Download individual files directly from the browser |
| 📦 Folder download | Download an entire folder (or the whole container) as a ZIP file |
| ⬆️ Upload | Upload files and entire folders with per-file progress bars |
| 🖱️ Drag & drop | Drop files onto the page to upload |
| ✅ Overwrite toggle | Checkbox in the upload panel — skip files that already exist, or overwrite them |
| ➕ New item | Create a new empty folder or a new blank file directly in the browser |
| ✏️ Rename | Rename files and folders in-place (contributors only) |
| 🗑️ Delete | Delete individual files or entire folder trees (contributors only) |
| ℹ️ Location info | View recursive stats for any folder: total sub-folders, files, and combined size |
| 🔑 SAS generator | Generate a User Delegation SAS URL for any file or folder — configurable expiry, IP restriction, and permissions (Contributors only) |
| 🔤 Sortable columns | Click any column header in list view (Name, Size, Modified, Created) to sort; click again to reverse |
| 🔍 Search | Search bar with two scopes: **Current folder** filters the visible items in real time; **Whole container** scans all blobs across the entire container |
| 👁️ File viewer | Preview text-based files (`.txt`, `.json`, `.log`, `.xml`, `.csv`, `.md`, etc.) directly in the browser |
| 📝 File editor | Edit the content of any viewable file directly in the browser — changes are saved back to Blob Storage (Contributors only) |
| 🏷️ Audit metadata | Each uploaded file stores `uploaded_by_upn` and `uploaded_by_oid` as blob metadata; overwrites add `last_edited_by_upn` and `last_edited_by_oid`. The uploader / last editor is shown as a subtitle under the filename in list and search views |
| 🔒 Upload RBAC | Upload, Rename, and Delete buttons appear automatically when the user has **Storage Blob Data Contributor** — no config needed |
| 🏷️ Permission badge | Header shows **✏️ Writer** or **📖 Reader** based on the signed-in user's detected role |
| 🗄️ Storage picker | Browse all Azure subscriptions, storage accounts, and containers accessible to the user via the ARM API, and switch targets without re-deploying |
| 🔗 Copy URL | Copy the direct blob URL to the clipboard with one click |
| 📄 File icons | Automatic icons based on file extension |
| 👤 User display | Signed-in user shown as `Display Name (upn@domain)` in the header |
| 📱 Responsive | Works on desktop and mobile |

---

## Architecture

```
Browser
  │
  ├─ index.html     — HTML shell (sign-in page, app layout, all modals)
  ├─ config.js      — Settings (client ID, tenant ID, storage account, feature flags)
  ├─ auth.js        — OAuth 2.0 PKCE flow (fetch + crypto.subtle)
  ├─ arm.js         — Azure Resource Manager API (subscription/account/container discovery)
  ├─ storage.js     — Azure Blob Storage REST API (list, download, upload, rename, delete, SAS)
  ├─ app.js         — UI logic (rendering, navigation, modals, upload queue, SAS modal)
  ├─ style.css      — Responsive stylesheet
  └─ favicon.svg    — App icon
```

All API calls go directly from the browser to the Azure REST APIs using bearer tokens obtained from Entra ID. No proxy or back-end is involved.

Two separate tokens are acquired:

| Token scope | Used for |
|---|---|
| `https://storage.azure.com/user_impersonation` | All Azure Blob Storage operations (list, upload, download, SAS key, etc.) |
| `https://management.azure.com/user_impersonation` | ARM — discovering subscriptions, storage accounts, and containers for the storage picker |

---

## Prerequisites

- An **Azure subscription**
- An **Azure Storage Account** with at least one blob container
- Permission to create an **App Registration** in Microsoft Entra ID

---

## Setup (step by step)

### 1 — Register an Entra ID App Registration

1. Open [Azure Portal](https://portal.azure.com) → **Microsoft Entra ID** → **App registrations** → **New registration**.
2. Give it a name (e.g. `Blob Explorer`).
3. Set **Supported account types** to match your organisation (typically *Single tenant*).
4. Under **Redirect URI**, choose platform **Single-page application (SPA)** and enter:
   - **Production:** `https://<storage-account>.z6.web.core.windows.net` *(your Static Website endpoint)*
   - **Local dev:** `http://localhost:3000`
   > Additional redirect URIs can be added later under **Authentication**.
5. Click **Register**.
6. Copy the **Application (client) ID** and **Directory (tenant) ID** — you will need them in `config.js`.

---

### 2 — Grant API permissions

1. In the App Registration → **API permissions** → **Add a permission**.
2. Choose **Azure Storage** → **Delegated permissions** → tick `user_impersonation` → **Add permissions**.
3. Add a second permission: choose **Azure Service Management** → **Delegated permissions** → tick `user_impersonation` → **Add permissions**.
   > The Azure Service Management permission is required for the **Storage Picker** feature. If you set fixed `accountName`/`containerName` in `config.js` and disable `allowStoragePicker`, this permission is not needed.
4. If your tenant requires it, click **Grant admin consent for \<tenant\>**.

---

### 3 — Assign Storage RBAC roles to users

Roles are assigned on the **Storage Account** (or scoped to a specific container) via **Access control (IAM)**.

| Role | Effect |
|---|---|
| **Storage Blob Data Reader** | Can sign in, browse, download, view properties, view location info, copy URLs, use the file viewer, and search. |
| **Storage Blob Data Contributor** | Everything above, **plus** Upload, Rename, Delete, New item, **Edit file**, and **SAS generation** become available. |

**Steps:**
1. Go to your **Storage Account** → **Access control (IAM)** → **Add role assignment**.
2. Select the desired role.
3. Assign it to the relevant users or security groups.

> The app detects the role automatically at login. The **✏️ Writer / 📖 Reader** badge in the header reflects the detected role.

---

### 4 — Configure CORS on the target Storage Account(s)

CORS must be enabled so the browser can call the Storage REST API from your Static Website origin.

> **Important:** CORS must be configured on every storage account you want to **browse**, not on the account that hosts the static website files. If you use the **Storage Picker** to switch between multiple accounts, each target account needs its own CORS rule.
>
> Example: if the app is deployed to `myapphost.z6.web.core.windows.net` and you browse blobs on `mydata`, you must add the CORS rule to **`mydata`** — not `myapphost`.

1. Go to the **target Storage Account** (the one you want to browse) → **Resource sharing (CORS)** → **Blob service**.
2. Add a rule:

   | Field | Value |
   |---|---|
   | Allowed origins | `https://<host-account>.z6.web.core.windows.net` *(the URL where the app is hosted)* |
   | Allowed methods | `GET, PUT, DELETE, OPTIONS, HEAD, POST` |
   | Allowed headers | `*` |
   | Exposed headers | `*` |
   | Max age | `86400` |

3. For **local development**, add a second rule to the target account:

   | Field | Value |
   |---|---|
   | Allowed origins | `http://localhost:3000` |
   | Allowed methods | `GET, PUT, DELETE, OPTIONS, HEAD, POST` |
   | Allowed headers | `*` |
   | Exposed headers | `*` |
   | Max age | `3600` |

> **Method reference:**
> - `GET` — download and blob listing
> - `HEAD` — properties panel and overwrite existence check
> - `PUT` — upload, rename (copy), and the RBAC probe
> - `DELETE` — delete blobs and the RBAC probe cleanup
> - `POST` — **required for SAS generation** (fetches a User Delegation Key from `?comp=userdelegationkey`)
> - `OPTIONS` — browser preflight

---

### 5 — Enable Static Website hosting

1. Go to your **Storage Account** → **Static website** → **Enabled**.
2. Set **Index document name** to `index.html`.
3. Set **Error document path** to `index.html`.
4. Note the **Primary endpoint** URL — this is your app's public URL.

---

### 6 — Update `config.js`

Open `config.js` and fill in your values:

```js
const CONFIG = {
  auth: {
    clientId:    "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",  // App Registration client ID
    tenantId:    "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",  // Entra ID tenant ID
    redirectUri: "https://<account>.z6.web.core.windows.net",
  },
  storage: {
    accountName:   "",   // Storage account name (no .blob.core.windows.net). Empty = always show picker.
    containerName: "",   // Container to browse. Empty = always show picker.
  },
  app: {
    title:              "Blob Storage Explorer",
    allowDownload:      true,   // Show download button on files
    allowRename:        true,   // Show rename option (contributors only)
    allowDelete:        true,   // Show delete option (contributors only)
    allowStoragePicker: true,   // Show "Change" button to switch storage account/container via ARM
    allowSas:           true,   // Show 🔑 SAS button on files and folders
    sasShowPermissions: false,  // false = hide permission checkboxes, force read-only SAS
  },
  upload: {
    maxFileSizeMB: 0,   // Per-file size limit in MB. 0 = no limit.
  },
};
```

#### Config flags reference

| Flag | Default | Description |
|---|---|---|
| `auth.clientId` | — | **Required.** App Registration client ID from Entra ID. |
| `auth.tenantId` | — | **Required.** Entra ID directory (tenant) ID. |
| `auth.redirectUri` | — | **Required.** Must match a registered SPA Redirect URI exactly. |
| `storage.accountName` | `""` | Storage account name. Leave empty to always prompt the picker. |
| `storage.containerName` | `""` | Container name. Leave empty to always prompt the picker. |
| `app.title` | `"Blob Storage Explorer"` | Title shown in the browser tab and app header. |
| `app.allowDownload` | `true` | Show or hide the download button on files. |
| `app.allowRename` | `true` | Show or hide rename buttons (only visible to Contributors anyway). |
| `app.allowDelete` | `true` | Show or hide delete buttons (only visible to Contributors anyway). |
| `app.allowStoragePicker` | `true` | Show or hide the **Change** button that opens the ARM-based storage/container picker. |
| `app.allowSas` | `true` | Show or hide the 🔑 **SAS** button on files and folders. |
| `app.sasShowPermissions` | `false` | `true` = show permission checkboxes in the SAS modal. `false` = hide them and always generate a read-only SAS. |
| `upload.maxFileSizeMB` | `0` | Maximum size for a single uploaded file in MB. `0` means no limit. |

---

### 7 — Deploy to the `$web` container

Upload all files to the `$web` blob container that Azure creates automatically when Static Website hosting is enabled.

**Using Azure CLI:**

```powershell
az storage blob upload-batch `
  --account-name <storage-account> `
  --destination '$web' `
  --source "C:\path\to\blob-explorer" `
  --overwrite
```

**Using Azure Storage Explorer (GUI):**
Open Storage Explorer → navigate to `Blob Containers` → `$web` → upload all files from the project folder.

Then open the **Primary endpoint** URL in your browser.

---

## Local development

Serve the files with any static HTTP server. The redirect URI `http://localhost:3000` must be registered in the App Registration.

```powershell
# Python
python -m http.server 3000

# Node.js
npx serve . -p 3000
```

Then open `http://localhost:3000`.

---

## How authentication works

The app implements the **OAuth 2.0 Authorization Code Flow with PKCE** entirely in the browser using only native APIs (`fetch`, `crypto.subtle`, `sessionStorage`).

```
1. User clicks "Sign in with Microsoft"
      ↓
2. App generates a PKCE code_verifier + code_challenge (SHA-256)
      ↓
3. Browser redirects to login.microsoftonline.com/authorize
      ↓
4. User authenticates with Microsoft
      ↓
5. Microsoft redirects back with ?code=...&state=...
      ↓
6. App calls /token endpoint via fetch → receives access_token + refresh_token
      ↓
7. Tokens stored in sessionStorage — page loads normally
      ↓
8. On subsequent visits: silent refresh via refresh_token (no redirect needed)
```

No MSAL or any other authentication library is used.

---

## How upload permissions work

Upload access (and Rename/Delete access) is determined entirely by Azure RBAC — no lists of users need to be configured in `config.js`.

On every login the app performs a **silent probe**: it attempts to `PUT` a zero-byte blob, then immediately `DELETE`s it.

- **201 Created** → user has `Storage Blob Data Contributor` → Upload, Rename, and Delete buttons appear. Header shows **✏️ Writer**.
- **403 Forbidden** → user has `Reader` only → those buttons stay hidden. Header shows **📖 Reader**.

The real enforcement is always done server-side by Azure; the probe just controls the UI.

---

## How SAS generation works

The **🔑 SAS** button (available on both files and folders when `allowSas: true`) generates a **User Delegation SAS** — it uses the signed-in user's OAuth token instead of the storage account key.

### Process

```
1. App calls POST https://<account>.blob.core.windows.net/
        ?restype=service&comp=userdelegationkey
   with the signed-in user's bearer token and a requested key window
      ↓
2. Azure returns a User Delegation Key (valid up to 7 days)
      ↓
3. App builds the canonical string-to-sign (Azure Blob SAS v2020-12-06)
      ↓
4. App signs with HMAC-SHA256 via browser crypto.subtle
      ↓
5. SAS URL is displayed in the modal — copy to clipboard with one click
```

### SAS options

| Field | Description |
|---|---|
| **Start time** | Optional. When the SAS becomes valid. Leave blank to start immediately. |
| **Expiry time** | Required. When the SAS expires. Maximum 7 days from now (User Delegation Key limit). |
| **Allowed IPs** | Optional. Restrict the SAS to a single IP (`1.2.3.4`) or an IP range (`1.2.3.4-1.2.3.9`). Leave blank to allow any IP. |
| **Permissions** | Shown only when `sasShowPermissions: true`. For files: Read / Write / Delete. For folders: Read / List / Write / Delete. |

### SAS scope

| Target | SAS type | Notes |
|---|---|---|
| Single file | Blob SAS (`sr=b`) | URL points directly to the file |
| Folder | Container SAS (`sr=c`) with path prefix | Grants access to all blobs under the folder prefix |

### Permission defaults

When `sasShowPermissions: false` (the default), the app always generates a **read-only** SAS:
- Files: `r` (read)
- Folders: `rl` (read + list)

When `sasShowPermissions: true`, the user can tick additional permissions (write, delete, list) in the modal.

### Requirements

- `POST` must be in the Blob service **CORS allowed methods** on the target account (see step 4).
- The signed-in user must have the **Storage Blob Data Contributor** role. The SAS button is hidden for Reader-only users.

---

## How rename works

Rename uses the Azure Blob Storage **Copy Blob** (`x-ms-copy-source`) + **Delete Blob** REST API sequence — there is no native server-side rename operation.

For **files**: one copy + one delete.

For **folders**: the app lists all blobs under the source prefix and renames each one individually. This is an O(n) operation — for a folder with 500 files, 1000 REST calls are made (500 copies + 500 deletes). Use with care on very large folders.

---

## Upload: overwrite behaviour

The upload panel has an **Overwrite existing files** checkbox:

| Checkbox | Behaviour |
|---|---|
| ✅ Checked (default) | Files are uploaded regardless of whether they already exist — existing blobs are replaced. |
| ☐ Unchecked | Before each upload the app sends a `HEAD` request to check whether the blob exists. Files that already exist are marked **Skipped** and no data is transferred. |

---

## Project files

| File | Purpose |
|---|---|
| `index.html` | HTML shell — sign-in page, main app layout, upload panel, all modals (properties, rename, new item, info, SAS) |
| `config.js` | ⚙️ **Edit this** — Entra ID client/tenant IDs, storage account, container, and all feature flags |
| `auth.js` | OAuth 2.0 PKCE flow — sign-in redirect, token exchange, silent refresh, sign-out |
| `arm.js` | Azure Resource Manager API — discover subscriptions, storage accounts, and containers for the storage picker |
| `storage.js` | Azure Blob Storage REST API — list (with metadata), download, upload (block blobs + custom metadata headers), rename (copy+delete), delete, properties (HEAD), RBAC probe, SAS generation |
| `app.js` | UI logic — rendering, navigation, view modes, sortable columns, search (folder + container-wide), upload queue, folder tree, file viewer, file editor, all modals, SAS modal, toast notifications |
| `style.css` | Responsive stylesheet |
| `favicon.svg` | SVG app icon (Azure-blue cloud) |

---

## How upload metadata works

Every blob written by the app carries up to four custom metadata fields (stored as `x-ms-meta-*` headers):

| Metadata key | Set on | Value |
|---|---|---|
| `uploaded_by_upn` | First upload | UPN of the user who originally uploaded the file |
| `uploaded_by_oid` | First upload | Entra ID object ID of the original uploader |
| `last_edited_by_upn` | Overwrite **or** in-browser Edit save | UPN of the user who last modified the file |
| `last_edited_by_oid` | Overwrite **or** in-browser Edit save | Entra ID object ID of the last editor |

The values are visible as a subtitle under each filename in **list view** and **search results** (e.g. `⬆ alice@contoso.com · ✏ bob@contoso.com`).

> Files uploaded before this feature was added will have no metadata fields and show no subtitle. The fields are populated only when the app performs the upload or edit — metadata is not back-filled automatically.

---

## Troubleshooting

| Symptom | Likely cause | Fix |
|---|---|---|
| `AADSTS50011` — Reply URL mismatch | `redirectUri` in `config.js` does not match a registered SPA Redirect URI | Add the exact current URL as a **Single-page application** Redirect URI in the App Registration |
| `AADSTS700054` — response_type not enabled | Wrong platform type selected | Delete the Web platform entry and re-add as **Single-page application** |
| 403 on file listing | User lacks a Storage data-plane role | Assign **Storage Blob Data Reader** via IAM on the storage account |
| 403 on upload | User lacks Contributor role | Assign **Storage Blob Data Contributor** via IAM |
| Properties panel shows "Failed to load" | CORS rule missing `HEAD` or `Exposed headers` not set to `*` | Update the CORS rule on the Blob service (see step 4) |
| Rename fails with copy error | User lacks Contributor role, or CORS missing `PUT`/`DELETE` | Assign **Storage Blob Data Contributor** and confirm CORS allows `PUT` and `DELETE` |
| SAS button returns "403 CORS not enabled or no matching rule found" | `POST` method is missing from the CORS rule | Add `POST` to **Allowed methods** in the CORS rule on the Blob service (see step 4) |
| SAS button returns 403 (not a CORS error) | User lacks Contributor role | Assign **Storage Blob Data Contributor** via IAM — SAS generation requires the Contributor role |
| SAS URL returns 403 after generation | IP restriction field set incorrectly | Check the IP/range field in the SAS modal, or leave it blank to allow any IP |
| SAS URL returns 403 after generation | SAS has expired or not yet started | Check the Start/Expiry times in the SAS modal |
| Storage picker shows no subscriptions | ARM permission not granted | Add **Azure Service Management → user_impersonation** in API permissions and grant admin consent |
| Storage picker shows no storage accounts | User lacks ARM Reader role | Assign the **Reader** role on the subscription or resource group via IAM |
| "Copy URL" button does nothing | Browser blocked clipboard access | Serve the app over HTTPS; clipboard API requires a secure context |
| CORS error in console | CORS rule missing, wrong origin, or missing methods | Update the CORS rule on the Blob service (see step 4) |
| Blank page after deployment | Wrong index document | Set **Index document** to `index.html` in Static Website settings |
| Upload button never appears | RBAC probe blocked | Confirm the user has **Storage Blob Data Contributor** and CORS allows `PUT` and `DELETE` |
| Files skipped unexpectedly | Overwrite checkbox is unchecked | Tick **Overwrite existing files** in the upload panel, then re-queue |
| Edit button not visible | User has Reader role, or file type is not viewable | Assign **Storage Blob Data Contributor** via IAM; the Edit button only appears for text-based file types |
| Edit save returns 403 | User lacks Contributor role | Assign **Storage Blob Data Contributor** via IAM |
| Audit metadata not shown under filename | File was uploaded before metadata tracking was added, or by a different tool | The subtitle only appears when `uploaded_by_upn` or `last_edited_by_upn` metadata is present on the blob |
| Token not refreshing | `offline_access` scope not consented | Sign out and sign in again to re-consent all scopes |
