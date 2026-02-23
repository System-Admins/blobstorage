// ============================================================
//  STORAGE — Azure Blob Storage REST API wrapper
//  Depends on: config.js, auth.js
// ============================================================

const _API_VERSION = "2020-10-02";

// ── SAS mode state ───────────────────────────────────────────
// When active, all API calls use the SAS token instead of Bearer auth.
const _SAS_STATE = {
  active:        false,   // true when browsing via a SAS URL
  accountName:   "",
  containerName: "",
  blobPrefix:    "",      // initial blob/folder prefix from the URL ("" = container root)
  sasQuery:      "",      // full query string including "?" — appended to every request
  permissions:   "",      // sp= value (e.g. "rl", "rwdl")
  expiry:        null,    // Date object for se= (null if not present)
  start:         null,    // Date object for st= (null if not present)
  signedResource:"",      // sr= value: "c" (container), "b" (blob), "d" (directory)
};

/** Returns true when the explorer is operating in SAS-token mode. */
function isSasMode() { return _SAS_STATE.active; }

/** Returns the current SAS mode state (read-only snapshot). */
function getSasState() { return { ..._SAS_STATE }; }

/**
 * Parse a SAS URL and activate SAS mode.
 * Supports container, folder (directory/prefix) and single-blob SAS URLs.
 *
 * Examples:
 *   https://account.blob.core.windows.net/container?sv=...&sig=...
 *   https://account.blob.core.windows.net/container/folder/path?sv=...&sig=...
 *   https://account.blob.core.windows.net/container/file.txt?sv=...&sig=...
 *
 * @param {string} sasUrl  Full SAS URL
 * @returns {{ accountName: string, containerName: string, blobPrefix: string, permissions: string }}
 */
function activateSasMode(sasUrl) {
  const info = parseSasUrl(sasUrl);

  _SAS_STATE.active         = true;
  _SAS_STATE.accountName    = info.accountName;
  _SAS_STATE.containerName  = info.containerName;
  _SAS_STATE.blobPrefix     = info.blobPrefix;
  _SAS_STATE.sasQuery       = info.sasQuery;
  _SAS_STATE.permissions    = info.permissions;
  _SAS_STATE.expiry         = info.expiry;
  _SAS_STATE.start          = info.start;
  _SAS_STATE.signedResource = info.signedResource;

  // Also update CONFIG.storage so the rest of the app sees the right account/container
  CONFIG.storage.accountName  = info.accountName;
  CONFIG.storage.containerName = info.containerName;

  return { accountName: info.accountName, containerName: info.containerName, blobPrefix: info.blobPrefix, permissions: info.permissions };
}

/**
 * Parse a SAS URL without mutating any global state.
 * Used for preview/validation before committing to SAS mode.
 *
 * @param {string} sasUrl  Full SAS URL
 * @returns {{ accountName: string, containerName: string, blobPrefix: string, permissions: string, sasQuery: string, expiry: Date|null, start: Date|null, signedResource: string }}
 */
function parseSasUrl(sasUrl) {
  const parsed = new URL(sasUrl);

  // Extract account name from hostname (e.g. "myaccount.blob.core.windows.net")
  const hostParts = parsed.hostname.split(".");
  if (hostParts.length < 4 || hostParts[1] !== "blob") {
    throw new Error("Invalid SAS URL: hostname must be <account>.blob.core.windows.net");
  }
  const accountName = hostParts[0];

  // Path: /<container>  or  /<container>/<blobPath>
  const pathSegments = parsed.pathname.split("/").filter(Boolean).map(decodeURIComponent);
  if (pathSegments.length === 0) {
    throw new Error("Invalid SAS URL: missing container name in the path.");
  }
  const containerName = pathSegments[0];
  const blobPrefix    = pathSegments.slice(1).join("/");

  // Validate that the URL has a signature
  if (!parsed.searchParams.get("sig")) {
    throw new Error("Invalid SAS URL: missing \"sig\" parameter. Please provide a complete SAS URL.");
  }

  // Extract SAS metadata
  const sp = parsed.searchParams.get("sp") || "";
  const se = parsed.searchParams.get("se") || "";
  const st = parsed.searchParams.get("st") || "";
  const sr = parsed.searchParams.get("sr") || "";

  return {
    accountName,
    containerName,
    blobPrefix,
    sasQuery:       parsed.search, // includes leading "?"
    permissions:    sp,
    expiry:         se ? new Date(se) : null,
    start:          st ? new Date(st) : null,
    signedResource: sr,
  };
}

/** Deactivate SAS mode (return to normal Bearer token auth). */
function deactivateSasMode() {
  _SAS_STATE.active         = false;
  _SAS_STATE.accountName    = "";
  _SAS_STATE.containerName  = "";
  _SAS_STATE.blobPrefix     = "";
  _SAS_STATE.sasQuery       = "";
  _SAS_STATE.permissions    = "";
  _SAS_STATE.expiry         = null;
  _SAS_STATE.start          = null;
  _SAS_STATE.signedResource = "";
}

/**
 * Build auth headers for a storage request.
 * In SAS mode returns empty object (token is in the query string).
 * In normal mode returns the Authorization + x-ms-version headers.
 */
async function _storageAuthHeaders() {
  if (_SAS_STATE.active) {
    return { "x-ms-version": _API_VERSION };
  }
  const token = await getStorageToken();
  return {
    Authorization:  `Bearer ${token}`,
    "x-ms-version": _API_VERSION,
  };
}

/**
 * Append the SAS query string to a URL when in SAS mode.
 * In normal mode returns the URL unchanged.
 * Handles URLs that may already have query parameters.
 */
function _sasUrl(url) {
  if (!_SAS_STATE.active) return url;
  const sep = url.includes("?") ? "&" : "?";
  // Strip the leading "?" from sasQuery when appending with "&"
  const qs = _SAS_STATE.sasQuery.startsWith("?") ? _SAS_STATE.sasQuery.slice(1) : _SAS_STATE.sasQuery;
  return url + sep + qs;
}

// ── Public API ───────────────────────────────────────────────

/**
 * List all virtual folders (BlobPrefix) and files (Blob) directly
 * under the given prefix, using "/" as delimiter.
 *
 * Automatically handles multi-page responses via NextMarker.
 *
 * @param {string} prefix  Folder path ending with "/" (empty string = root)
 * @returns {Promise<{folders: FolderItem[], files: FileItem[]}>}
 */
async function listBlobsAtPrefix(prefix = "") {
  const { accountName, containerName } = CONFIG.storage;

  let allFolders = [];
  let allFiles   = [];
  let marker     = null;

  do {
    const authHeaders = await _storageAuthHeaders();
    const url      = _sasUrl(_buildListUrl(accountName, containerName, prefix, marker));
    const response = await fetch(url, {
      headers: {
        ...authHeaders,
        "x-ms-date":    new Date().toUTCString(),
      },
    });

    if (!response.ok) {
      const text = await response.text();
      if (response.status === 403) {
        throw new Error(_SAS_STATE.active
          ? "Access denied (403). The SAS token does not include 'List' (l) permission. Regenerate the SAS URL with list permission enabled."
          : "Access denied (403). The signed-in user does not have a Storage data-plane role on this container. " +
            "Assign the 'Storage Blob Data Reader' role via Access control (IAM) on the storage account, then sign out and back in."
        );
      }
      throw new Error(`Storage API ${response.status}: ${_parseStorageError(text)}`);
    }

    const { folders, files, nextMarker } = _parseListXml(
      await response.text(),
      prefix
    );

    allFolders = allFolders.concat(folders);
    allFiles   = allFiles.concat(files);
    marker     = nextMarker;

  } while (marker);

  return { folders: allFolders, files: allFiles };
}

/**
 * Download a blob via the REST API (bearer-token auth) and trigger
 * a browser save-file dialog.
 *
 * @param {string} blobName  Full blob path inside the container
 */
async function downloadBlob(blobName) {
  const { accountName, containerName } = CONFIG.storage;
  const authHeaders = await _storageAuthHeaders();

  const url = _sasUrl(`https://${accountName}.blob.core.windows.net`
            + `/${containerName}/${_encodePath(blobName)}`);

  const response = await fetch(url, {
    headers: authHeaders,
  });

  if (!response.ok) {
    if (response.status === 403) {
      throw new Error(_SAS_STATE.active
        ? "Access denied (403). The SAS token does not include 'Read' (r) permission. Regenerate the SAS URL with read permission enabled."
        : "Access denied (403). Assign the 'Storage Blob Data Reader' role via IAM on the storage account."
      );
    }
    throw new Error(`Download failed (${response.status}): ${response.statusText}`);
  }

  const blob     = await response.blob();
  const fileName = blobName.split("/").pop();

  // Trigger the browser download dialog
  const objectUrl = URL.createObjectURL(blob);
  const anchor    = document.createElement("a");
  anchor.href     = objectUrl;
  anchor.download = fileName;
  document.body.appendChild(anchor);
  anchor.click();
  document.body.removeChild(anchor);
  URL.revokeObjectURL(objectUrl);
}

// ── Upload ────────────────────────────────────────────────────

const _CHUNK_SIZE = 4 * 1024 * 1024; // 4 MB per block

/**
 * Upload a browser File object to Azure Blob Storage.
 * Files ≤ 4 MB are sent in a single PUT request.
 * Larger files are uploaded as committed block blobs (Azure block upload API).
 *
 * Requires the signed-in user to have the
 * "Storage Blob Data Contributor" role on the container.
 *
 * @param {string}        blobPath    Full blob name inside the container
 * @param {File}          file        Browser File object
 * @param {function|null} onProgress  Called with an integer 0-100 during upload
 * @param {object}        metadata    Optional key/value pairs stored as blob metadata
 */
async function uploadBlob(blobPath, file, onProgress, metadata = {}) {
  const { accountName, containerName } = CONFIG.storage;
  const contentType = file.type || "application/octet-stream";
  const blobUrl     = `https://${accountName}.blob.core.windows.net`
                    + `/${containerName}/${_encodePath(blobPath)}`;

  onProgress?.(0);

  if (file.size <= _CHUNK_SIZE) {
    // ── Single PUT (small file) ───────────────────────────────────────
    await _putWholeBlob(blobUrl, file, contentType, metadata);
    onProgress?.(100);
    return;
  }

  // ── Block-blob upload (large file) ──────────────────────────
  const blockCount = Math.ceil(file.size / _CHUNK_SIZE);
  const blockIds   = [];

  for (let i = 0; i < blockCount; i++) {
    const start   = i * _CHUNK_SIZE;
    const chunk   = file.slice(start, Math.min(start + _CHUNK_SIZE, file.size));
    const blockId = btoa(String(i).padStart(10, "0")); // fixed-length base64 block ID
    blockIds.push(blockId);

    const authHeaders = await _storageAuthHeaders();
    const blockUrl = _sasUrl(`${blobUrl}?comp=block&blockid=${encodeURIComponent(blockId)}`);

    const res = await fetch(blockUrl, {
      method: "PUT",
      headers: authHeaders,
      body: chunk,
    });

    if (!res.ok) {
      const text = await res.text();
      if (res.status === 403) throw new Error(_SAS_STATE.active
        ? "Upload denied (403). The SAS token does not include write permission (w/c/a). Regenerate the SAS URL with write permission enabled."
        : "Upload denied (403). Assign the 'Storage Blob Data Contributor' role via IAM on the storage account."
      );
      throw new Error(`Block upload failed (${res.status}): ${_parseStorageError(text)}`);
    }

    // Report progress up to 95% — last 5% is the commit
    onProgress?.(Math.round(((i + 1) / blockCount) * 95));
  }

  // ── Commit block list ────────────────────────────────────────
  const commitAuthHeaders = await _storageAuthHeaders();
  const blockListXml =
    `<?xml version="1.0" encoding="utf-8"?><BlockList>` +
    blockIds.map(id => `<Latest>${id}</Latest>`).join("") +
    `</BlockList>`;

  const commitRes = await fetch(_sasUrl(`${blobUrl}?comp=blocklist`), {
    method: "PUT",
    headers: {
      ...commitAuthHeaders,
      "Content-Type":             "application/xml",
      "x-ms-blob-content-type":   contentType,
      ..._buildMetaHeaders(metadata),
    },
    body: blockListXml,
  });

  if (!commitRes.ok) {
    const text = await commitRes.text();
    throw new Error(`Block list commit failed (${commitRes.status}): ${_parseStorageError(text)}`);
  }

  onProgress?.(100);
}

async function _putWholeBlob(blobUrl, file, contentType, metadata = {}) {
  const authHeaders = await _storageAuthHeaders();
  const res   = await fetch(_sasUrl(blobUrl), {
    method:  "PUT",
    headers: {
      ...authHeaders,
      "x-ms-blob-type": "BlockBlob",
      "Content-Type":  contentType,
      "Content-Length": String(file.size),
      ..._buildMetaHeaders(metadata),
    },
    body: file,
  });
  if (!res.ok) {
    const text = await res.text();
    if (res.status === 403) throw new Error(_SAS_STATE.active
      ? "Upload denied (403). The SAS token does not include write permission (w/c/a). Regenerate the SAS URL with write permission enabled."
      : "Upload denied (403). Assign the 'Storage Blob Data Contributor' role via IAM on the storage account."
    );
    throw new Error(`Upload failed (${res.status}): ${_parseStorageError(text)}`);
  }
}

/**
 * Build a flat object of x-ms-meta-* headers from a metadata key/value map.
 * Skips any entry with an empty key or value.
 */
function _buildMetaHeaders(metadata = {}) {
  const headers = {};
  for (const [k, v] of Object.entries(metadata)) {
    if (k && v !== "" && v !== null && v !== undefined) {
      headers[`x-ms-meta-${k}`] = String(v);
    }
  }
  return headers;
}

/**
 * Return the custom metadata for a blob as a plain { key: value } object,
 * with the "x-ms-meta-" prefix stripped from each key.
 * @param {string} blobName  Full blob path inside the container
 */
async function getBlobMetadata(blobName) {
  const props = await getBlobProperties(blobName);
  const meta  = {};
  for (const [key, value] of Object.entries(props)) {
    if (key.startsWith("x-ms-meta-")) meta[key.slice(10)] = value;
  }
  return meta;
}

/**
 * Silently probe whether the signed-in user has write access to the container.
 * Uploads a zero-byte marker blob and immediately deletes it.
 * Returns true  → user has Storage Blob Data Contributor (or higher).
 * Returns false → user has Reader only (or the probe failed).
 */
async function probeUploadPermission() {
  // In SAS mode, derive permission from the sp= parameter — no probe needed
  if (_SAS_STATE.active) {
    // Write access requires at least 'w' (write) or 'c' (create) or 'a' (add)
    return /[wca]/.test(_SAS_STATE.permissions);
  }

  const { accountName, containerName } = CONFIG.storage;
  let token;
  try { token = await getStorageToken(); } catch { return false; }

  const probeUrl = `https://${accountName}.blob.core.windows.net`
                 + `/${containerName}/.upload-probe-${Date.now()}`;

  // Attempt a zero-byte PUT.
  // NOTE: If the user only has Reader rights, Azure Storage returns 403
  // WITHOUT an Access-Control-Allow-Origin header. The browser enforces CORS
  // and blocks the response, causing fetch() to throw a TypeError rather than
  // returning a Response object. This CORS warning in the browser console is
  // expected and unavoidable — it is Azure's intentional behaviour for
  // unauthorised requests. We catch the TypeError here and treat it as "no
  // write permission", which is the correct outcome.
  let putRes;
  try {
    putRes = await fetch(probeUrl, {
      method: "PUT",
      headers: {
        Authorization:     `Bearer ${token}`,
        "x-ms-version":   _API_VERSION,
        "x-ms-blob-type": "BlockBlob",
      },
      body: "",
    });
  } catch {
    // TypeError thrown by the browser due to missing CORS headers on the 403 —
    // user does not have write access.
    return false;
  }

  if (!putRes.ok) return false; // 403 with CORS headers present, or other error

  // Clean up — fire-and-forget DELETE (ignore errors)
  fetch(probeUrl, {
    method:  "DELETE",
    headers: {
      Authorization:   `Bearer ${token}`,
      "x-ms-version": _API_VERSION,
    },
  }).catch(() => {});

  return true;
}

/**
 * Delete a single blob.
 * @param {string} blobName  Full blob path inside the container
 */
async function deleteBlob(blobName) {
  const { accountName, containerName } = CONFIG.storage;
  const authHeaders = await _storageAuthHeaders();
  const url   = _sasUrl(`https://${accountName}.blob.core.windows.net`
              + `/${containerName}/${_encodePath(blobName)}`);

  const res = await fetch(url, {
    method:  "DELETE",
    headers: authHeaders,
  });

  if (!res.ok) {
    const text = await res.text();
    if (res.status === 403) throw new Error(_SAS_STATE.active
      ? "Delete denied (403). The SAS token does not include 'Delete' (d) permission. Regenerate the SAS URL with delete permission enabled."
      : "Delete denied (403). Assign the 'Storage Blob Data Contributor' role via IAM on the storage account."
    );
    throw new Error(`Delete failed (${res.status}): ${_parseStorageError(text)}`);
  }
}

// ── Append Blob helpers (used by audit log) ──────────────────

/**
 * Create an Append Blob if it doesn't exist, then append a block of text.
 * Uses "If-None-Match: *" for creation (no-op if blob exists) and
 * ?comp=appendblock for appending.
 *
 * @param {string} blobPath  Full blob path inside the container
 * @param {string} text      UTF-8 text to append
 */
async function appendToBlob(blobPath, text) {
  const { accountName, containerName } = CONFIG.storage;
  const blobUrl = `https://${accountName}.blob.core.windows.net`
                + `/${containerName}/${_encodePath(blobPath)}`;

  // 1. Ensure the append blob exists (idempotent create)
  const authCreate = await _storageAuthHeaders();
  const createRes = await fetch(_sasUrl(blobUrl), {
    method: "PUT",
    headers: {
      ...authCreate,
      "x-ms-blob-type":    "AppendBlob",
      "Content-Length":     "0",
      "Content-Type":      "application/x-ndjson",
      "If-None-Match":     "*",          // only create if it doesn't exist
    },
  });
  // 201 = created, 409/412 = already exists (or precondition failed) — all are fine
  if (!createRes.ok && createRes.status !== 409 && createRes.status !== 412) {
    const text2 = await createRes.text();
    throw new Error(`Append blob create failed (${createRes.status}): ${_parseStorageError(text2)}`);
  }

  // 2. Append the data block
  const body = new TextEncoder().encode(text);
  const authAppend = await _storageAuthHeaders();
  const appendRes = await fetch(_sasUrl(`${blobUrl}?comp=appendblock`), {
    method: "PUT",
    headers: {
      ...authAppend,
      "Content-Length": String(body.byteLength),
      "Content-Type":  "application/x-ndjson",
    },
    body: body,
  });
  if (!appendRes.ok) {
    const text3 = await appendRes.text();
    throw new Error(`Append block failed (${appendRes.status}): ${_parseStorageError(text3)}`);
  }
}

/**
 * Read a blob's content as text (UTF-8).
 * Returns the text string, or null if the blob does not exist (404).
 * @param {string} blobPath  Full blob path inside the container
 * @returns {Promise<string|null>}
 */
async function readBlobText(blobPath) {
  const { accountName, containerName } = CONFIG.storage;
  const authHeaders = await _storageAuthHeaders();

  const url = _sasUrl(`https://${accountName}.blob.core.windows.net`
            + `/${containerName}/${_encodePath(blobPath)}`);

  const response = await fetch(url, { headers: authHeaders });

  if (response.status === 404) return null;
  if (!response.ok) {
    const text = await response.text();
    throw new Error(`Read blob failed (${response.status}): ${_parseStorageError(text)}`);
  }
  return response.text();
}

/**
 * Delete all blobs under a virtual folder prefix (recursive).
 * Deletes files in parallel batches for performance on large folders.
 * @param {string} prefix  Folder path ending with "/"
 */
async function deleteFolderContents(prefix) {
  // Recursively collect ALL blob names first (avoids interleaving list + delete)
  async function collectAll(pfx) {
    const { folders, files } = await listBlobsAtPrefix(pfx);
    let names = files.map(f => f.name);
    // Recurse subfolders in parallel
    const sub = await Promise.all(folders.map(f => collectAll(f.name)));
    for (const s of sub) names = names.concat(s);
    return names;
  }

  const allNames = await collectAll(prefix);
  const BATCH = 50;
  for (let i = 0; i < allNames.length; i += BATCH) {
    await Promise.all(allNames.slice(i, i + BATCH).map(n => deleteBlob(n)));
  }
}

/**
 * Download all blobs under a virtual folder prefix as a single ZIP file.
 * Uses "store" (no-compression) mode so no external library is required.
 * @param {string}   prefix       Folder path ending with "/" (e.g. "reports/")
 * @param {string}   displayName  Base name for the .zip file
 * @param {Function} [onProgress] Called as onProgress(fetchedCount, totalCount) after each file
 */
async function downloadFolderAsZip(prefix, displayName, onProgress) {
  const { accountName, containerName } = CONFIG.storage;

  // Recursively collect every blob name under this prefix (parallel recursion)
  async function collectFiles(pfx) {
    const { folders, files } = await listBlobsAtPrefix(pfx);
    let names = files.map((f) => f.name);
    const sub = await Promise.all(folders.map(f => collectFiles(f.name)));
    for (const s of sub) names = names.concat(s);
    return names;
  }

  const allNames = await collectFiles(prefix);
  if (allNames.length === 0) throw new Error("Folder is empty — nothing to download.");

  // Fetch blobs in parallel batches for speed
  const FETCH_BATCH = 6;
  const entries = [];
  let fetched = 0;
  for (let i = 0; i < allNames.length; i += FETCH_BATCH) {
    const batch = allNames.slice(i, i + FETCH_BATCH);
    const results = await Promise.all(batch.map(async (name) => {
      const authHeaders = await _storageAuthHeaders();
      const url = _sasUrl(`https://${accountName}.blob.core.windows.net/${containerName}/${_encodePath(name)}`);
      const res = await fetch(url, { headers: authHeaders });
      if (!res.ok) throw new Error(`Failed to fetch "${name}" (${res.status})`);
      const data = new Uint8Array(await res.arrayBuffer());
      const relativeName = name.startsWith(prefix) ? name.slice(prefix.length) : name;
      return { name: relativeName, data };
    }));
    entries.push(...results);
    fetched += results.length;
    if (typeof onProgress === "function") onProgress(fetched, allNames.length);
  }

  const zipBlob = _buildZip(entries);
  const objUrl  = URL.createObjectURL(zipBlob);
  const anchor  = document.createElement("a");
  anchor.href     = objUrl;
  anchor.download = `${displayName || "folder"}.zip`;
  document.body.appendChild(anchor);
  anchor.click();
  document.body.removeChild(anchor);
  URL.revokeObjectURL(objUrl);
}

// ── Private helpers ──────────────────────────────────────────

/**
 * Fetch all properties and metadata for a single blob via a HEAD request.
 * @param {string} blobName  Full blob path inside the container
 * @returns {Promise<object>} Key/value map of headers
 */
async function getBlobProperties(blobName) {
  const { accountName, containerName } = CONFIG.storage;
  const authHeaders = await _storageAuthHeaders();
  const url   = _sasUrl(`https://${accountName}.blob.core.windows.net`
              + `/${containerName}/${_encodePath(blobName)}`);

  const res = await fetch(url, {
    method:  "HEAD",
    headers: authHeaders,
  });

  if (!res.ok) throw new Error(`Properties request failed (${res.status})`);

  // Collect all response headers into a plain object
  const props = {};
  res.headers.forEach((value, key) => { props[key] = value; });
  return props;
}

/**
 * Check whether a blob exists (HEAD request).
 * @returns {Promise<boolean>}
 */
async function blobExists(blobName) {
  const { accountName, containerName } = CONFIG.storage;
  const authHeaders = await _storageAuthHeaders();
  const url   = _sasUrl(`https://${accountName}.blob.core.windows.net`
              + `/${containerName}/${_encodePath(blobName)}`);
  const res = await fetch(url, {
    method:  "HEAD",
    headers: authHeaders,
  });
  return res.status === 200;
}

/**
 * Copy a blob to a new destination (no delete of source).
 * @param {string} srcName   Source blob path
 * @param {string} destName  Destination blob path
 */
async function copyBlobOnly(srcName, destName) {
  const { accountName, containerName } = CONFIG.storage;
  const authHeaders = await _storageAuthHeaders();
  const baseUrl  = `https://${accountName}.blob.core.windows.net/${containerName}`;
  const srcUrl   = `${baseUrl}/${_encodePath(srcName)}`;
  const destUrl  = `${baseUrl}/${_encodePath(destName)}`;

  const copyRes = await fetch(_sasUrl(destUrl), {
    method:  "PUT",
    headers: {
      ...authHeaders,
      "x-ms-copy-source": _sasUrl(srcUrl),
    },
  });
  if (!copyRes.ok) {
    const text = await copyRes.text();
    throw new Error(`Copy failed (${copyRes.status}): ${_parseStorageError(text)}`);
  }
}

/**
 * Rename a blob by copying it to the new name then deleting the source.
 * Works for both files and virtual folder prefixes (copies all blobs below).
 * @param {string} srcName   Source blob path
 * @param {string} destName  Destination blob path
 */
async function renameBlob(srcName, destName) {
  try { await copyBlobOnly(srcName, destName); }
  catch (err) { throw new Error(`Rename (copy) failed: ${err.message}`); }
  try { await deleteBlob(srcName); }
  catch (err) { throw new Error(`Rename (delete source) failed: ${err.message}`); }
}

function _buildListUrl(accountName, containerName, prefix, marker) {
  const params = new URLSearchParams({
    restype:   "container",
    comp:      "list",
    delimiter: "/",
    include:   "metadata",
  });
  if (prefix) params.set("prefix", prefix);
  if (marker) params.set("marker", marker);
  return `https://${accountName}.blob.core.windows.net/${containerName}?${params}`;
}

/**
 * Parse a single <Blob> XML element into a file metadata object.
 * Shared by _parseListXml() and listAllBlobs() to avoid duplicating the parsing logic.
 *
 * @param {Element} blob         A <Blob> DOM element from the Storage REST API XML response
 * @param {string}  displayName  Display name to use (prefix-stripped or full path)
 * @returns {object}  Parsed file metadata
 */
function _parseBlobElement(blob, displayName) {
  const name  = blob.querySelector("Name")?.textContent ?? "";
  const props = blob.querySelector("Properties");
  const meta  = blob.querySelector("Metadata");
  const metaVal = (key) => meta?.querySelector(key)?.textContent ?? "";
  return {
    name,
    displayName,
    type:            "file",
    size:            parseInt(props?.querySelector("Content-Length")?.textContent ?? "0", 10),
    lastModified:    props?.querySelector("Last-Modified")?.textContent ?? "",
    createdOn:       props?.querySelector("Creation-Time")?.textContent  ?? "",
    contentType:     props?.querySelector("Content-Type")?.textContent  ?? "application/octet-stream",
    etag:            props?.querySelector("Etag")?.textContent ?? "",
    md5:             props?.querySelector("Content-MD5")?.textContent ?? "",
    uploadedByUpn:   metaVal("uploaded_by_upn"),
    uploadedByOid:   metaVal("uploaded_by_oid"),
    lastEditedByUpn: metaVal("last_edited_by_upn"),
    lastEditedByOid: metaVal("last_edited_by_oid"),
  };
}

function _parseListXml(xmlText, currentPrefix) {
  const doc = new DOMParser().parseFromString(xmlText, "application/xml");

  // Virtual directories (folders)
  const folders = [...doc.querySelectorAll("BlobPrefix Name")].map((el) => {
    const fullName   = el.textContent;
    const displayName = fullName.slice(currentPrefix.length).replace(/\/$/, "");
    return { name: fullName, displayName, type: "folder" };
  });

  // Individual blobs (files)
  const files = [...doc.querySelectorAll("Blobs > Blob")].map((blob) => {
    const name = blob.querySelector("Name")?.textContent ?? "";
    return _parseBlobElement(blob, name.slice(currentPrefix.length));
  });

  const rawMarker = doc.querySelector("NextMarker")?.textContent ?? "";
  return { folders, files, nextMarker: rawMarker || null };
}

function _parseStorageError(xmlText) {
  try {
    const doc = new DOMParser().parseFromString(xmlText, "application/xml");
    return doc.querySelector("Message")?.textContent ?? xmlText;
  } catch {
    return xmlText;
  }
}

/** Encode each path segment individually so "/" separators are preserved. */
function _encodePath(path) {
  return path.split("/").map(encodeURIComponent).join("/");
}

/**
 * Build a ZIP blob from an array of {name, data} entries.
 *
 * Compression method: STORE (method 0 — no compression).
 * Deflate compression is intentionally omitted to keep this implementation
 * dependency-free and universally compatible. The files being packaged are
 * already downloaded from Blob Storage (often already compressed formats such
 * as images, videos, or Office documents), so a second compression pass would
 * yield minimal size savings while adding significant CPU overhead.
 *
 * @param {{name: string, data: Uint8Array}[]} entries
 * @returns {Blob}
 */
function _buildZip(entries) {
  if (entries.length > 65535) {
    throw new Error(
      `Too many files for ZIP format (${entries.length} entries, max 65\u202F535). ` +
      `Download smaller folders individually.`
    );
  }
  const enc         = new TextEncoder();
  const localParts  = [];
  const centralDir  = [];
  let   offset      = 0;

  for (const entry of entries) {
    const nameBytes = enc.encode(entry.name);
    const data      = entry.data;
    const crc       = _crc32(data);
    const size      = data.length;

    // Local file header (30 bytes + filename)
    const lfh = new Uint8Array(30 + nameBytes.length);
    const lfv = new DataView(lfh.buffer);
    lfv.setUint32( 0, 0x04034b50, true); // signature
    lfv.setUint16( 4, 20,         true); // version needed
    lfv.setUint16( 6,  0x0800,    true); // flags (bit 11 = UTF-8 filename encoding)
    lfv.setUint16( 8,  0,         true); // compression (store)
    lfv.setUint16(10,  0,         true); // last mod time
    lfv.setUint16(12,  0,         true); // last mod date
    lfv.setUint32(14, crc,        true); // crc-32
    lfv.setUint32(18, size,       true); // compressed size
    lfv.setUint32(22, size,       true); // uncompressed size
    lfv.setUint16(26, nameBytes.length, true); // filename length
    lfv.setUint16(28,  0,         true); // extra field length
    lfh.set(nameBytes, 30);

    // Central directory entry (46 bytes + filename)
    const cde = new Uint8Array(46 + nameBytes.length);
    const cdv = new DataView(cde.buffer);
    cdv.setUint32( 0, 0x02014b50, true); // signature
    cdv.setUint16( 4, 20,         true); // version made by
    cdv.setUint16( 6, 20,         true); // version needed
    cdv.setUint16( 8,  0x0800,    true); // flags (bit 11 = UTF-8 filename encoding)
    cdv.setUint16(10,  0,         true); // compression
    cdv.setUint16(12,  0,         true); // last mod time
    cdv.setUint16(14,  0,         true); // last mod date
    cdv.setUint32(16, crc,        true); // crc-32
    cdv.setUint32(20, size,       true); // compressed size
    cdv.setUint32(24, size,       true); // uncompressed size
    cdv.setUint16(28, nameBytes.length, true); // filename length
    cdv.setUint16(30,  0,         true); // extra field length
    cdv.setUint16(32,  0,         true); // file comment length
    cdv.setUint16(34,  0,         true); // disk number start
    cdv.setUint16(36,  0,         true); // internal attributes
    cdv.setUint32(38,  0,         true); // external attributes
    cdv.setUint32(42, offset,     true); // offset of local header
    cde.set(nameBytes, 46);

    localParts.push(lfh, data);
    centralDir.push(cde);
    offset += lfh.length + size;
  }

  const cdSize = centralDir.reduce((s, b) => s + b.length, 0);

  // End of central directory record (22 bytes)
  const eocd = new Uint8Array(22);
  const eov  = new DataView(eocd.buffer);
  eov.setUint32( 0, 0x06054b50,      true); // signature
  eov.setUint16( 4,  0,              true); // disk number
  eov.setUint16( 6,  0,              true); // disk with CD
  eov.setUint16( 8, entries.length,  true); // entries on disk
  eov.setUint16(10, entries.length,  true); // total entries
  eov.setUint32(12, cdSize,          true); // CD size
  eov.setUint32(16, offset,          true); // CD offset
  eov.setUint16(20,  0,              true); // comment length

  return new Blob([...localParts, ...centralDir, eocd], { type: "application/zip" });
}

/** CRC-32 lookup table (pre-computed for the standard ZIP polynomial 0xEDB88320). */
const _CRC32_TABLE = new Uint32Array(256);
for (let n = 0; n < 256; n++) {
  let c = n;
  for (let k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
  _CRC32_TABLE[n] = c;
}

/** CRC-32 checksum using a pre-computed lookup table (10-50× faster than bit-shifting per byte). */
function _crc32(data) {
  let crc = 0xFFFFFFFF;
  for (let i = 0; i < data.length; i++) {
    crc = _CRC32_TABLE[(crc ^ data[i]) & 0xFF] ^ (crc >>> 8);
  }
  return (crc ^ 0xFFFFFFFF) >>> 0;
}

// ── Shared formatting utilities (used by app.js) ─────────────

function formatFileSize(bytes) {
  if (bytes === 0) return "0 B";
  const units = ["B", "KB", "MB", "GB", "TB"];
  const i = Math.floor(Math.log(bytes) / Math.log(1024));
  return `${(bytes / Math.pow(1024, i)).toFixed(i === 0 ? 0 : 1)} ${units[i]}`;
}

function formatDate(dateStr) {
  if (!dateStr) return "—";
  return new Date(dateStr).toLocaleString();
}

function formatDateShort(dateStr) {
  if (!dateStr) return "—";
  return new Date(dateStr).toLocaleDateString();
}

function getFileIcon(name) {
  const ext = (name.split(".").pop() ?? "").toLowerCase();
  const map = {
    // Images
    jpg: "🖼️", jpeg: "🖼️", png: "🖼️", gif: "🖼️", svg: "🖼️", webp: "🖼️", bmp: "🖼️",
    // Documents
    pdf: "📄", doc: "📝", docx: "📝", xls: "📊", xlsx: "📊", ppt: "📊", pptx: "📊",
    txt: "📃", md: "📃", csv: "📊",
    // Data / config
    json: "📋", xml: "📋", yaml: "📋", yml: "📋", toml: "📋", ini: "📋",
    // Code
    js: "💻", ts: "💻", html: "💻", css: "💻", py: "💻", java: "💻", cs: "💻",
    cpp: "💻", c: "💻", go: "💻", rs: "💻", sh: "💻", ps1: "💻", rb: "💻",
    // Archives
    zip: "🗜️", tar: "🗜️", gz: "🗜️", rar: "🗜️", "7z": "🗜️",
    // Video / audio
    mp4: "🎬", avi: "🎬", mov: "🎬", mkv: "🎬", mp3: "🎵", wav: "🎵", flac: "🎵",
    // Executables
    exe: "⚙️", dll: "⚙️", msi: "⚙️",
  };
  return map[ext] ?? "📄";
}

/**
 * Recursively counts all subfolders, files and total byte size under a prefix.
 * @param {string} prefix  Folder path ending with "/"
 * @returns {Promise<{totalFolders: number, totalFiles: number, totalSize: number}>}
 */
async function getFolderStats(prefix) {
  let totalFolders = 0;
  let totalFiles   = 0;
  let totalSize    = 0;

  async function recurse(pfx) {
    const { folders, files } = await listBlobsAtPrefix(pfx);
    totalFolders += folders.length;
    totalFiles   += files.length;
    totalSize    += files.reduce((sum, f) => sum + (f.size || 0), 0);
    // Recurse into all subfolders in parallel
    await Promise.all(folders.map(sub => recurse(sub.name)));
  }

  await recurse(prefix);
  return { totalFolders, totalFiles, totalSize };
}

// ── User Delegation SAS ──────────────────────────────────────

/**
 * Generate a User Delegation SAS URL for a single blob or an entire container.
 * Requires the signed-in user to have the Storage Blob Data Contributor role.
 *
 * @param {string}  blobName   Full blob path inside the container (pass "" for container SAS)
 * @param {boolean} isFolder   true  → container-level SAS (sr=c)
 *                             false → blob-level SAS       (sr=b)
 * @param {{ start: string, expiry: string, permissions: string, ip: string }} options
 *   start       – ISO 8601 UTC, e.g. "2024-01-01T00:00:00Z" (empty = omit)
 *   expiry      – ISO 8601 UTC (required)
 *   permissions – canonical-order chars, e.g. "rl", "rwdl"
 *   ip          – optional single IP or range, e.g. "10.0.0.1-10.0.0.255"
 * @returns {Promise<string>} Complete SAS URL
 */
async function generateSasToken(blobName, isFolder, options) {
  const { accountName, containerName } = CONFIG.storage;
  const token = await getStorageToken();
  const sv    = "2020-12-06"; // UDK / SAS API version

  // ── 1. Obtain a User Delegation Key ─────────────────────────
  // Key window: (now − 5 min) → min(SAS expiry, now + 7 days)
  const keyStart  = _toIso8601(new Date(Date.now() - 5 * 60 * 1000));
  const seDate    = new Date(options.expiry);
  const maxExpiry = new Date(Date.now() + 7 * 24 * 60 * 60 * 1000);
  const keyExpiry = _toIso8601(seDate > maxExpiry ? maxExpiry : seDate);

  const keyBody = `<?xml version="1.0" encoding="utf-8"?><KeyInfo>`
                + `<Start>${keyStart}</Start>`
                + `<Expiry>${keyExpiry}</Expiry></KeyInfo>`;

  const keyRes = await fetch(
    `https://${accountName}.blob.core.windows.net/?restype=service&comp=userdelegationkey`,
    {
      method:  "POST",
      headers: {
        Authorization:  `Bearer ${token}`,
        "x-ms-version": sv,
        "x-ms-date":    new Date().toUTCString(),
        "Content-Type": "application/xml",
      },
      body: keyBody,
    }
  );
  if (!keyRes.ok) {
    const text = await keyRes.text();
    throw new Error(
      `User delegation key failed (${keyRes.status}): ${_parseStorageError(text)}`
    );
  }

  const keyDoc = new DOMParser().parseFromString(await keyRes.text(), "application/xml");
  const skoid  = keyDoc.querySelector("SignedOid")?.textContent     ?? "";
  const sktid  = keyDoc.querySelector("SignedTid")?.textContent     ?? "";
  const skt    = keyDoc.querySelector("SignedStart")?.textContent   ?? "";
  const ske    = keyDoc.querySelector("SignedExpiry")?.textContent  ?? "";
  const sks    = keyDoc.querySelector("SignedService")?.textContent ?? "b";
  const skv    = keyDoc.querySelector("SignedVersion")?.textContent ?? sv;
  const rawKey = keyDoc.querySelector("Value")?.textContent         ?? "";

  // ── 2. Build the string-to-sign (format: API version 2020-12-06) ──
  const sp  = options.permissions;
  const st  = options.start  || "";
  const se  = options.expiry;
  const sip = options.ip     || "";
  const spr = "https";
  const sr  = isFolder ? "c" : "b";

  const canonicalizedResource = isFolder
    ? `/blob/${accountName}/${containerName}`
    : `/blob/${accountName}/${containerName}/${blobName}`;

  const stringToSign = [
    sp,                    // signedPermissions
    st,                    // signedStart
    se,                    // signedExpiry
    canonicalizedResource, // canonicalizedResource
    skoid,                 // signedKeyObjectId
    sktid,                 // signedKeyTenantId
    skt,                   // signedKeyStart
    ske,                   // signedKeyExpiry
    sks,                   // signedKeyService
    skv,                   // signedKeyVersion
    "",                    // signedAuthorizedUserObjectId
    "",                    // signedUnauthorizedUserObjectId
    "",                    // signedCorrelationId
    sip,                   // signedIP
    spr,                   // signedProtocol
    sv,                    // signedVersion
    sr,                    // signedResource
    "",                    // signedSnapshotTime
    "",                    // signedEncryptionScope
    "",                    // rscc  (response-cache-control)
    "",                    // rscd  (response-content-disposition)
    "",                    // rsce  (response-content-encoding)
    "",                    // rscl  (response-content-language)
    "",                    // rsct  (response-content-type)
  ].join("\n");

  // ── 3. HMAC-SHA256 signature via Web Crypto API ──────────────
  const keyBytes  = Uint8Array.from(atob(rawKey), c => c.charCodeAt(0));
  const cryptoKey = await crypto.subtle.importKey(
    "raw", keyBytes, { name: "HMAC", hash: "SHA-256" }, false, ["sign"]
  );
  const sigBytes = await crypto.subtle.sign(
    "HMAC", cryptoKey, new TextEncoder().encode(stringToSign)
  );
  const sig = btoa(String.fromCharCode(...new Uint8Array(sigBytes)));

  // ── 4. Assemble the SAS URL ───────────────────────────────────
  const qs = new URLSearchParams();
  qs.set("sv",    sv);
  qs.set("sr",    sr);
  qs.set("sp",    sp);
  qs.set("se",    se);
  if (st)  qs.set("st",   st);
  if (sip) qs.set("sip",  sip);
  qs.set("spr",   spr);
  qs.set("skoid", skoid);
  qs.set("sktid", sktid);
  qs.set("skt",   skt);
  qs.set("ske",   ske);
  qs.set("sks",   sks);
  qs.set("skv",   skv);
  qs.set("sig",   sig);

  const baseUrl = isFolder
    ? `https://${accountName}.blob.core.windows.net/${containerName}`
    : `https://${accountName}.blob.core.windows.net/${containerName}/${_encodePath(blobName)}`;

  return `${baseUrl}?${qs.toString()}`;
}
function _toIso8601(date) {
  return date.toISOString().replace(/\.\d{3}Z$/, "Z");
}

/**
 * List every blob in the container (no delimiter — flat list of all files)
 * and optionally filter to blobs whose name contains `nameFilter`.
 * Used for whole-container search.
 *
 * @param {string} [nameFilter]  Case-insensitive substring to match against the full blob path
 * @returns {Promise<FileItem[]>}
 */
async function listAllBlobs(nameFilter = "") {
  const { accountName, containerName } = CONFIG.storage;
  const lowerFilter = nameFilter.toLowerCase();
  let allFiles = [];
  const folderSet = new Set();
  let marker   = null;

  do {
    const authHeaders = await _storageAuthHeaders();
    const params = new URLSearchParams({ restype: "container", comp: "list", include: "metadata" });
    if (marker) params.set("marker", marker);
    const url = _sasUrl(`https://${accountName}.blob.core.windows.net/${containerName}?${params}`);

    const response = await fetch(url, {
      headers: {
        ...authHeaders,
        "x-ms-date":    new Date().toUTCString(),
      },
    });

    if (!response.ok) {
      const text = await response.text();
      throw new Error(`Storage API ${response.status}: ${_parseStorageError(text)}`);
    }

    const doc = new DOMParser().parseFromString(await response.text(), "application/xml");

    const blobs = [...doc.querySelectorAll("Blobs > Blob")].map((blob) =>
      _parseBlobElement(blob, blob.querySelector("Name")?.textContent ?? "")
    );

    // Collect all virtual folder prefixes from every blob
    for (const b of blobs) {
      const parts = b.name.split("/");
      for (let i = 1; i < parts.length; i++) {
        folderSet.add(parts.slice(0, i).join("/") + "/");
      }
    }

    // Exclude virtual-directory placeholder blobs; apply name filter
    const filtered = blobs.filter((b) => {
      if (b.name.endsWith("/.keep")) return false;
      if (b.name === ".audit" || b.name.startsWith(".audit/")) return false;
      return !lowerFilter || b.name.toLowerCase().includes(lowerFilter);
    });

    allFiles = allFiles.concat(filtered);
    marker   = doc.querySelector("NextMarker")?.textContent || null;
    if (marker === "") marker = null;
  } while (marker);

  // Build matching folder objects
  const matchingFolders = [...folderSet]
    .filter(f => !f.startsWith(".audit/"))
    .filter(f => !lowerFilter || f.toLowerCase().includes(lowerFilter))
    .map(name => ({
      name,
      displayName: name.slice(0, -1).split("/").pop(),
      type: "folder",
    }));

  return { files: allFiles, folders: matchingFolders };
}
