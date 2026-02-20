// ============================================================
//  CONFIGURATION — update before deploying
// ============================================================
const CONFIG = {

  // ── Entra ID (Azure AD) ───────────────────────────────────
  auth: {
    // Application (client) ID from your App Registration
    clientId: "CLIENT ID GOES HERE",

    // Directory (tenant) ID — or "common" for multi-tenant
    tenantId: "TENANT ID GOES HERE",

    // Must match a Redirect URI registered as a "Single-page application"
    // Local dev  → http://localhost:3000
    // Production → https://<account>.z13.web.core.windows.net
    redirectUri: 'https://contoso.z6.web.core.windows.net',
  },

  // ── Azure Blob Storage ────────────────────────────────────
  storage: {
    // Storage account name (without .blob.core.windows.net).
    // Leave empty ("") to always show the storage picker after sign-in.
    // When set, the picker is skipped and this account/container is opened
    // directly — users can still switch via the "Change" button in the header.
    accountName: "",

    // Blob container to browse. Leave empty to always show the picker.
    containerName: "",
  },

  // ── App behaviour ─────────────────────────────────────────
  app: {
    title: "Blob Storage Explorer",

    // Show download button for files and folders
    // Requires "Storage Blob Data Reader" (or higher) role on the container
    allowDownload: true,

    // Show rename button (also requires Contributor role at runtime)
    allowRename: true,

    // Show delete button (also requires Contributor role at runtime)
    allowDelete: true,

    // Show the "Change" button in the header and allow users to switch to a
    // different storage account or container after sign-in.
    // Set to false to lock users to the account/container defined above.
    allowStoragePicker: true,

    // Allow generating User Delegation SAS tokens for files and folders.
    // Requires Storage Blob Data Contributor (or higher) on the storage account.
    allowSas: true,

    // Show the Permissions checkboxes in the SAS modal.
    // Set to false to hide them and always issue a read-only SAS.
    sasShowPermissions: false,
  },

  // ── Upload ─────────────────────────────────────────────────
  // Upload availability is detected automatically at runtime by probing
  // whether the signed-in user has the "Storage Blob Data Contributor"
  // role on the container. No manual user/group lists are needed.
  upload: {
    // Maximum individual file size in MB (0 = no limit).
    maxFileSizeMB: 0,
  },
};
