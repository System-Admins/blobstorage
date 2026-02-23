# Blob Browser

A zero-dependency, static single-page web application for browsing and managing files in **Azure Blob Storage**. Deploy it as an Azure Storage Account **Static Website** — no back-end, no build step, no `node_modules`. Authenticates via **OAuth 2.0 PKCE** against **Microsoft Entra ID** (or via a pasted **SAS URL**) using only native browser APIs.

<p align="center">
  <img src="docs/demo.gif" alt="Blob Browser demo" width="800" />
</p>

---

## Features

### File management
- **Browse & navigate** — virtual folder hierarchy, breadcrumbs, and a collapsible folder-tree sidebar
- **Upload** — single files, multiple files, or entire folders via drag-and-drop with progress bars and overwrite-conflict handling
- **Download** — individual files or entire folders as client-side ZIP archives
- **Create / Rename / Delete / Copy / Move** — full CRUD on blobs and virtual folders (Contributor role required)
- **View & Edit** — preview images, view text/code with syntax line numbers, and edit text files in-browser

### Search
- **Instant filter** — type to filter the current folder in real time
- **Deep scan** — search across the entire container (prefix-based server-side listing)

### Reports & audit
- **Container report** — interactive modal with summary dashboard (file/folder counts, total size), filterable table with all 12 blob properties, and CSV download
- **Audit log** — automatic action logging (upload, delete, rename, move, copy, edit, etc.) stored as daily Append Blobs in `.audit/YYYY/MM/DD.jsonl`
- **Audit viewer** — built-in modal with date navigation, action and text filters, and CSV export
- **Audit metadata** — records the uploader and last editor on every blob via `x-ms-meta-*` headers

### Sharing & SAS
- **SAS generator** — create User Delegation SAS URLs with configurable expiry, IP restriction, and permissions
- **Copy URL** — copy blob URL or application deep-link to clipboard, with email integration
- **App links** — shareable deep-links that survive OAuth redirects
- **Email links** — send file/folder links via Microsoft Graph (`Mail.Send`) from the signed-in user's mailbox

### Navigation & UX
- **Storage picker** — browse subscriptions, storage accounts, and containers via the Azure Resource Manager API
- **Folder tree sidebar** — right-click context menu with New, Upload, Download, Copy URL, Info, Copy, Move, Rename, SAS, Delete, Refresh, and Report
- **Dark / light theme** — toggle manually or auto-detect from OS preference; persisted in `localStorage`
- **Keyboard shortcuts** — `Ctrl+K` search, `Ctrl+U` upload, `F5` refresh, `Backspace` go up, `?` help
- **Responsive** — works on desktop and mobile
- **SAS-only mode** — paste a SAS URL from the sign-in screen to browse without an Entra ID account

### Security & compliance
- **Auto RBAC detection** — UI adapts to Reader vs. Contributor role automatically; no configuration needed
- **Content Security Policy** — strict CSP meta tag; no inline scripts
- **Zero dependencies** — no third-party JavaScript libraries; everything runs on native `fetch`, `crypto.subtle`, and standard DOM APIs

---

## Prerequisites

| Requirement | Notes |
|---|---|
| **Azure subscription** | With at least one Storage Account and blob container |
| **Entra ID App Registration** | Permission to register an application (or an admin to do it for you) |
| **Azure CLI** *(optional)* | For the one-line deploy command below; you can also upload via the portal |

---

## Quick start

```
git clone <repo-url>
# Edit src/config.js with your clientId and tenantId (see Setup below)
az storage blob upload-batch --account-name <account> --destination '$web' --source ./src --overwrite
```

---

## Setup

### 1 — Create an App Registration

1. [Azure Portal](https://portal.azure.com) → **Microsoft Entra ID** → **App registrations** → **New registration**
2. Name it (e.g. `Blob Browser`), set **Supported account types** to **Single tenant**
3. **Redirect URI** → platform **Single-page application (SPA)** → enter your Static Website URL:
   ```
   https://<storage-account>.<zone>.web.core.windows.net
   ```
4. Click **Register** and copy the **Application (client) ID** and **Directory (tenant) ID**

### 2 — Add API permissions

| API | Type | Permission | Required for |
|---|---|---|---|
| **Azure Storage** | Delegated | `user_impersonation` | All blob operations |
| **Azure Service Management** | Delegated | `user_impersonation` | Storage Picker *(skip if using a fixed account/container)* |
| **Microsoft Graph** | Delegated | `Mail.Send` | Email integration *(optional)* |

Click **Grant admin consent** if required by your tenant.

### 3 — Assign Storage RBAC roles

On the **Storage Account** (or container) → **Access control (IAM)** → **Add role assignment**:

| Role | Capabilities |
|---|---|
| **Storage Blob Data Reader** | Browse, download, search, view, copy URL |
| **Storage Blob Data Contributor** | All of the above + upload, rename, delete, edit, create folders, generate SAS |

The app auto-detects the assigned role at runtime and shows a **Writer** or **Reader** badge accordingly.

### 4 — Configure CORS

On each **target** storage account you want to browse (not the one hosting the static site):

**Storage Account** → **Resource sharing (CORS)** → **Blob service** → add a rule:

| Field | Value |
|---|---|
| Allowed origins | `https://<host-account>.<zone>.web.core.windows.net` |
| Allowed methods | `GET, HEAD, PUT, DELETE, POST, OPTIONS` |
| Allowed headers | `*` |
| Exposed headers | `*` |
| Max age | `86400` |

> Add a second rule with origin `http://localhost:3000` for local development.

### 5 — Enable Static Website hosting

1. **Storage Account** → **Data management** → **Static website** → **Enabled**
2. Set **Index document** to `index.html` and **Error document** to `index.html`
3. Note the **Primary endpoint** URL

### 6 — Update `config.js`

Open `src/config.js` and fill in your App Registration details:

```js
const CONFIG = {
  auth: {
    clientId:  "YOUR-CLIENT-ID",     // Application (client) ID
    tenantId:  "YOUR-TENANT-ID",     // Directory (tenant) ID
    // redirectUri is auto-detected from window.location.origin
  },
  storage: {
    accountName:   "",               // Leave empty → show storage picker
    containerName: "",               // Leave empty → show storage picker
  },
  app: {
    title:              "Blob Browser",
    allowDownload:      true,
    allowRename:        true,
    allowDelete:        true,
    allowStoragePicker: true,
    allowSas:           true,
    sasShowPermissions: false,        // false → always generate read-only SAS
    allowEmail:         true,
  },
  upload: {
    maxFileSizeMB: 0,                // 0 = no limit
  },
};
```

All feature flags are optional. The app validates `clientId` and `tenantId` on load and will throw a clear error if placeholders are still present.

### 7 — Deploy

Upload the contents of `src/` to the `$web` blob container:

**PowerShell (Azure CLI):**
```powershell
az storage blob upload-batch `
  --account-name <storage-account> `
  --destination '$web' `
  --source ./src `
  --overwrite
```

**Bash (Azure CLI):**
```bash
az storage blob upload-batch \
  --account-name <storage-account> \
  --destination '$web' \
  --source ./src \
  --overwrite
```

Open the **Primary endpoint** URL in your browser.

---

## Project structure

```
src/
├── index.html      HTML shell and all modal markup
├── style.css       Responsive stylesheet with dark/light theme support
├── config.js       Settings — client/tenant IDs, storage target, feature flags
├── auth.js         OAuth 2.0 PKCE — sign-in, token exchange, silent refresh, Graph token
├── arm.js          ARM API — subscription/account/container discovery for the storage picker
├── storage.js      Blob Storage REST API — list, upload, download, rename, delete, SAS
├── app.js          UI — rendering, navigation, modals, upload queue, folder tree, search
└── favicon.svg     App icon
```

---

## How it works

| Aspect | Detail |
|---|---|
| **Authentication** | OAuth 2.0 Authorization Code with PKCE via `fetch` + `crypto.subtle`. Tokens stored in `sessionStorage`. No MSAL or third-party libraries. |
| **SAS-only mode** | Users can paste a SAS URL to browse without signing in. The UI adapts to the token's permissions (e.g. read-only if no write permission). |
| **Permission detection** | On sign-in the app probes write access by attempting a PUT + DELETE of a zero-byte sentinel blob (`.upload-probe-*`). Success → Contributor UI; failure → Reader UI. Enforcement is always server-side. |
| **SAS generation** | Fetches a User Delegation Key via `POST`, then signs the SAS string in-browser with HMAC-SHA256 (`crypto.subtle`). Maximum expiry: 7 days. |
| **Rename / Move** | Uses Copy Blob → Delete Source (no native rename in Azure). Folder renames are O(n) — one copy + delete per blob. |
| **ZIP download** | Folders are streamed blob-by-blob and zipped client-side before download. |
| **Audit logging** | Every mutating action (upload, delete, rename, move, copy, edit, create folder, SAS generation) is recorded as a JSONL entry in `.audit/YYYY/MM/DD.jsonl` using Append Blobs. The `.audit` folder is hidden from normal browsing. |
| **Audit metadata** | Every upload and edit stamps `x-ms-meta-uploadedBy` / `x-ms-meta-lastEditedBy` on the blob. |
| **Reports** | Container report recursively lists all blobs, displays an interactive summary (file/folder counts, total size), and offers CSV download with 12 columns (Type, Name, Full Path, Path Length, Blob URL, Size, Content Type, Last Modified, Created On, ETag, MD5, etc.). |
| **Deep links** | Shareable URLs with hash fragment `#a=<account>&c=<container>&p=<path>`. The hash is saved to `sessionStorage` before OAuth redirects and restored afterwards. |
| **Theme** | CSS custom properties with a `[data-theme="dark"]` selector on `<html>`. Follows OS preference by default; manual toggle persisted in `localStorage`. |

---

## Keyboard shortcuts

| Shortcut | Action |
|---|---|
| `Ctrl + K` | Toggle search bar |
| `Ctrl + U` | Toggle upload panel |
| `F5` | Refresh current folder (without reloading the page) |
| `Backspace` | Navigate to parent folder |
| `?` | Show help / keyboard shortcuts reference |

---

## Troubleshooting

| Symptom | Fix |
|---|---|
| `AADSTS50011` reply URL mismatch | Add the exact Static Website URL as a **SPA** Redirect URI in the App Registration |
| `AADSTS700054` response_type error | Delete the **Web** platform entry and re-add as **Single-page application** |
| 403 on listing / download | Assign **Storage Blob Data Reader** on the storage account or container |
| 403 on upload / rename / delete | Assign **Storage Blob Data Contributor** on the storage account or container |
| Upload button never appears | Confirm Contributor role and that CORS allows `PUT` and `DELETE` methods |
| SAS 403 / CORS error | Ensure `POST` is in CORS allowed methods on the **target** account |
| Storage picker shows nothing | Add **Azure Service Management** API permission + grant admin consent; assign **Reader** role on the subscription |
| CORS errors in console | Add/fix the CORS rule on the **target** storage account (not the hosting account) |
| Blank page after deploy | Set both index and error document to `index.html` in Static Website settings |
| Sign-in loop / no token | Clear `sessionStorage`, check that the redirect URI **exactly** matches (scheme, host, no trailing slash) |

---

## Browser support

Built on standard web APIs (`fetch`, `crypto.subtle`, `Blob`, `ReadableStream`, CSS custom properties). Works in all modern evergreen browsers:

- Chrome / Edge 90+
- Firefox 90+
- Safari 15+

---

## License

[MIT](LICENSE) — Copyright (c) 2026 System Admins
