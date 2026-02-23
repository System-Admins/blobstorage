// ============================================================
//  APP — Main application logic and UI rendering
//  Depends on: config.js, auth.js, arm.js, storage.js
// ============================================================

let _currentPrefix = "";
let _viewMode      = "list"; // "list" | "grid"
let _canUpload      = false;
let _listLoadFailed  = false;   // true when the initial blob listing fails (network/permissions)
let _listingPromise  = null;    // Promise for the in-flight _loadFiles call — probe waits on this
let _selection       = new Set(); // selected blob names

// ── Sort / Search state ───────────────────────────────────────
let _sortKey                = "name"; // "name" | "size" | "modified" | "created"
let _sortDir                = "asc";  // "asc" | "desc"
let _cachedFolders          = [];     // raw folders from last _loadFiles call
let _cachedFiles            = [];     // raw files from last _loadFiles call
let _containerSearchResults = null;   // null = normal view; { files, folders } = container-wide search hits
let _searchDebounceTimer    = null;

// sessionStorage key for persisting the user's storage selection
const _STORAGE_SELECTION_KEY = "be_storage_selection";

// ── Deep-link hash helpers ─────────────────────────────────────

/**
 * Build a hash fragment that encodes account, container, and blob path.
 * Format: #a=<account>&c=<container>&p=<path>
 */
function _buildAppHash(path) {
  const a = encodeURIComponent(CONFIG.storage.accountName);
  const c = encodeURIComponent(CONFIG.storage.containerName);
  const p = path ? encodeURIComponent(path) : "";
  return p ? `#a=${a}&c=${c}&p=${p}` : `#a=${a}&c=${c}`;
}

/**
 * Build a full app URL (origin + path + hash) for the given blob path.
 */
function _buildAppUrl(path) {
  return `${window.location.origin}${window.location.pathname}${_buildAppHash(path)}`;
}

/**
 * Parse the current URL hash. Supports both the new structured format
 * (#a=...&c=...&p=...) and the legacy format (#blobPath).
 * Returns { accountName, containerName, path }.
 */
function _parseAppHash() {
  const raw = window.location.hash.slice(1);
  if (!raw) return { accountName: "", containerName: "", path: "" };

  // New format: #a=<account>&c=<container>&p=<path>
  if (raw.startsWith("a=") || raw.includes("&c=")) {
    const params = new URLSearchParams(raw);
    return {
      accountName:   params.get("a") || "",
      containerName: params.get("c") || "",
      path:          params.get("p") || "",
    };
  }

  // Legacy format: #<blobPath>
  return { accountName: "", containerName: "", path: decodeURIComponent(raw) };
}

// Derived permission flags — depend on both RBAC probe and config switches
function _canRenameItems() { return _canUpload && (CONFIG.app.allowRename !== false); }
function _canDeleteItems() { return _canUpload && (CONFIG.app.allowDelete !== false); }
function _canSas()         { return !isSasMode() && _canUpload && CONFIG.app.allowSas !== false; }
function _canCopyItems()   { return _canUpload; }
function _canMoveItems()   { return _canUpload; }
function _canEditItems()   { return _canUpload; }
function _canEmail()       { return !isSasMode() && CONFIG.app.allowEmail !== false && !!getUser(); }

// ── Audit log ────────────────────────────────────────────────
// Writes JSONL entries to .audit/YYYY/MM/DD.jsonl (Append Blobs).
// Fire-and-forget: failures are silently logged to the console.

const _AUDIT_FOLDER = ".audit";

/**
 * Record an audit event.  Silently no-ops when the user lacks write access,
 * when in SAS mode, or when an error occurs (audit must never block the UI).
 *
 * @param {string} action   One of: download, upload, edit, delete, rename, copy, move, sas, create
 * @param {string} path     Blob or folder path the action targeted
 * @param {object} [details]  Optional extra context (destination, SAS expiry, etc.)
 */
function _audit(action, path, details) {
  // Guard: only audit when the user can write and is authenticated (not SAS mode)
  if (!_canUpload || isSasMode()) return;
  try {
    const user = getUser();
    const now  = new Date();
    const yyyy = now.getUTCFullYear();
    const mm   = String(now.getUTCMonth() + 1).padStart(2, "0");
    const dd   = String(now.getUTCDate()).padStart(2, "0");
    const logPath = `${_AUDIT_FOLDER}/${yyyy}/${mm}/${dd}.jsonl`;

    const entry = {
      ts:        now.toISOString(),
      user:      user?.username || user?.name || "unknown",
      userId:    user?.oid || "",
      action,
      path,
      ...(details && Object.keys(details).length ? { details } : {}),
    };

    // Fire-and-forget — do not await
    appendToBlob(logPath, JSON.stringify(entry) + "\n").catch((err) => {
      console.warn("[audit] Failed to write audit log:", err.message);
    });
  } catch (err) {
    console.warn("[audit] Error building audit entry:", err.message);
  }
}

// ── Audit log viewer ─────────────────────────────────────────

/**
 * Open the audit-log viewer modal.  Defaults to today's date.
 * Reads .audit/YYYY/MM/DD.jsonl and renders a filterable table.
 */
function _showAuditModal() {
  const modal   = _el("auditModal");
  const body    = _el("auditModalBody");
  const metaEl  = _el("auditMeta");
  const dateEl  = _el("auditDate");

  // Default to today (local time)
  const today = new Date();
  dateEl.value = _auditDateStr(today);

  modal.classList.remove("hidden");

  // State
  let _allEntries = [];
  let _filterAction = "";
  let _filterText   = "";

  // Close
  const close = () => { modal.classList.add("hidden"); body.innerHTML = ""; };
  _el("auditModalClose").onclick = close;
  _el("auditCloseBtn").onclick   = close;
  modal.onclick = (e) => { if (e.target === modal) close(); };

  // CSV export
  const csvBtn = _el("auditExportCsv");
  csvBtn.onclick = () => {
    if (!_allEntries.length) return;
    // Apply same filters as the table
    let entries = _allEntries;
    if (_filterAction) entries = entries.filter((e) => e.action === _filterAction);
    if (_filterText) {
      const lower = _filterText.toLowerCase();
      entries = entries.filter((e) =>
        (e.path || "").toLowerCase().includes(lower) ||
        (e.user || "").toLowerCase().includes(lower) ||
        JSON.stringify(e.details || {}).toLowerCase().includes(lower)
      );
    }
    const csvEsc = (v) => { const s = String(v ?? ""); return s.includes(",") || s.includes('"') || s.includes("\n") ? '"' + s.replace(/"/g, '""') + '"' : s; };
    const rows = ["Time,Action,User,Path,Details"];
    for (const e of entries) {
      const time = e.ts ? new Date(e.ts).toLocaleString() : "";
      const details = e.details ? Object.entries(e.details).map(([k, v]) => `${k}: ${v}`).join("; ") : "";
      rows.push([time, e.action || "", e.user || "", e.path || "", details].map(csvEsc).join(","));
    }
    const blob = new Blob([rows.join("\n")], { type: "text/csv" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = `audit-log-${dateEl.value}.csv`;
    a.click();
    URL.revokeObjectURL(a.href);
  };

  // Navigation
  _el("auditPrevDay").onclick = () => {
    const d = new Date(dateEl.value + "T00:00:00");
    d.setDate(d.getDate() - 1);
    dateEl.value = _auditDateStr(d);
    loadDay();
  };
  _el("auditNextDay").onclick = () => {
    const d = new Date(dateEl.value + "T00:00:00");
    d.setDate(d.getDate() + 1);
    dateEl.value = _auditDateStr(d);
    loadDay();
  };
  dateEl.onchange = loadDay;
  _el("auditRefresh").onclick = loadDay;

  async function loadDay() {
    const val = dateEl.value;           // "YYYY-MM-DD"
    if (!val) return;
    const [yyyy, mm, dd] = val.split("-");
    const logPath = `${_AUDIT_FOLDER}/${yyyy}/${mm}/${dd}.jsonl`;

    body.innerHTML = `<div class="audit-loading"><span class="audit-spinner"></span> Loading audit log\u2026</div>`;
    metaEl.textContent = "";

    try {
      const text = await readBlobText(logPath);
      if (text === null || text.trim() === "") {
        _allEntries = [];
        csvBtn.disabled = true;
        body.innerHTML = `<div class="audit-empty">No audit entries for ${_esc(val)}.</div>`;
        metaEl.textContent = "0 entries";
        return;
      }

      _allEntries = text.trim().split("\n").map((line, i) => {
        try { return JSON.parse(line); }
        catch { return { ts: "", user: "?", action: "?", path: `(parse error line ${i + 1})`, _raw: line }; }
      });

      csvBtn.disabled = _allEntries.length === 0;
      _filterAction = "";
      _filterText   = "";
      _renderAuditEntries();
    } catch (err) {
      body.innerHTML = `<div class="audit-empty">\u26A0\uFE0F Error loading audit log:<br><code>${_esc(err.message)}</code></div>`;
      metaEl.textContent = "";
    }
  }

  function _renderAuditEntries() {
    let entries = _allEntries;

    // Apply filters
    if (_filterAction) {
      entries = entries.filter((e) => e.action === _filterAction);
    }
    if (_filterText) {
      const lower = _filterText.toLowerCase();
      entries = entries.filter((e) =>
        (e.path || "").toLowerCase().includes(lower) ||
        (e.user || "").toLowerCase().includes(lower) ||
        JSON.stringify(e.details || {}).toLowerCase().includes(lower)
      );
    }

    // Collect unique actions for the filter dropdown
    const actions = [...new Set(_allEntries.map((e) => e.action).filter(Boolean))].sort();

    // Build filter bar + table
    let html = `<div class="audit-filter-bar">
      <select class="audit-filter-select" id="auditActionFilter" title="Filter by action">
        <option value="">All actions</option>
        ${actions.map((a) => `<option value="${_esc(a)}"${a === _filterAction ? " selected" : ""}>${_esc(a)}</option>`).join("")}
      </select>
      <input type="text" class="audit-filter-input" id="auditTextFilter" placeholder="Filter path, user, details\u2026" value="${_esc(_filterText)}" />
    </div>`;

    if (entries.length === 0) {
      html += `<div class="audit-empty">No entries match the current filter.</div>`;
    } else {
      html += `<table class="audit-table"><thead><tr>
        <th>Time</th>
        <th>Action</th>
        <th>User</th>
        <th>Path</th>
        <th>Details</th>
      </tr></thead><tbody>`;

      for (const e of entries) {
        const time = e.ts ? new Date(e.ts).toLocaleString() : "\u2014";
        const action = e.action || "\u2014";
        const actionClass = `audit-action-${_esc(action)}`;
        const user = _esc(e.user || "\u2014");
        const path = _esc(e.path || "\u2014");
        const details = e.details ? _esc(Object.entries(e.details).map(([k, v]) => `${k}: ${v}`).join(", ")) : "";

        html += `<tr>
          <td class="audit-col-time">${_esc(time)}</td>
          <td class="audit-col-action"><span class="audit-action-badge ${actionClass}">${_esc(action)}</span></td>
          <td class="audit-col-user" title="${user}">${user}</td>
          <td class="audit-col-path" title="${path}">${path}</td>
          <td class="audit-col-details" title="${details}">${details}</td>
        </tr>`;
      }
      html += `</tbody></table>`;
    }

    body.innerHTML = html;
    metaEl.textContent = `${entries.length} of ${_allEntries.length} entries`;

    // Wire filter controls
    const actionSel = _el("auditActionFilter");
    const textInp   = _el("auditTextFilter");
    if (actionSel) actionSel.onchange = () => { _filterAction = actionSel.value; _renderAuditEntries(); };
    if (textInp) {
      let _debounce;
      textInp.oninput = () => {
        clearTimeout(_debounce);
        _debounce = setTimeout(() => { _filterText = textInp.value; _renderAuditEntries(); }, 250);
      };
    }
  }

  // Load today on open
  loadDay();
}

/** Format a Date as "YYYY-MM-DD" for the date input. */
function _auditDateStr(d) {
  const yyyy = d.getFullYear();
  const mm   = String(d.getMonth() + 1).padStart(2, "0");
  const dd   = String(d.getDate()).padStart(2, "0");
  return `${yyyy}-${mm}-${dd}`;
}

// ── Entry point ──────────────────────────────────────────────

document.addEventListener("DOMContentLoaded", async () => {
  _showLoading(true);
  try {
    const account = await initAuth();
    if (account) {
      sessionStorage.removeItem("be_silent_tried"); // clean up if a silent attempt just succeeded
      // Restore a previously chosen storage selection (survives page refresh)
      _restoreStorageSelection();
      // If the URL contains a deep-link with account/container, apply it now
      // so the picker is skipped and the correct account is opened directly.
      const dl = _parseAppHash();
      if (dl.accountName && dl.containerName) {
        _setStorageSelection(dl.accountName, dl.containerName);
      }
      // Show picker when no storage is configured, unless the picker is disabled
      const pickerEnabled = CONFIG.app.allowStoragePicker !== false;
      if (pickerEnabled && (!CONFIG.storage.accountName || !CONFIG.storage.containerName)) {
        _showLoading(false);
        _showPickerPage(account);
      } else {
        _bootApp(account);
      }
    } else {
      // No stored session. On the very first load, attempt silent SSO (prompt=none)
      // so users with an active Entra ID / M365 session sign in automatically.
      // We mark sessionStorage *before* navigating so the return trip knows we
      // already tried — avoiding an infinite silent-redirect loop.
      const alreadyTriedSilent = sessionStorage.getItem("be_silent_tried");
      if (!alreadyTriedSilent) {
        sessionStorage.setItem("be_silent_tried", "1");
        _showLoading(true);
        await signInSilent();
        // signInSilent() navigates away — nothing below executes.
      } else {
        // Silent attempt already made and failed (no active SSO session).
        // Remove the flag and show the sign-in page.
        sessionStorage.removeItem("be_silent_tried");
        _showLoading(false);
        _showSignInPage();
      }
    }
  } catch (err) {
    console.error("[app] Init error:", err);
    _showLoading(false);
    _showSignInPage();
  }
});

// ── Sign-in / sign-out ───────────────────────────────────────

function _showSignInPage() {
  _el("signInPage").classList.remove("hidden");
  _el("mainApp").classList.add("hidden");
  _el("pickerPage").classList.add("hidden");
  _el("headerInfoBar").classList.add("hidden");
  document.title = CONFIG.app.title;

  _el("signInBtn").addEventListener("click", async () => {
    _showLoading(true);
    try {
      await signIn(); // Redirects to Microsoft login — page navigates away
    } catch (err) {
      console.error("[app] Sign-in error:", err);
      _showError("Sign-in failed. Please try again.");
      _showLoading(false);
    }
  }, { once: true });

  // Wire the "Open SAS" button on the sign-in page
  _el("openSasSignInBtn").addEventListener("click", () => _showOpenSasModal(), { once: false });
}

// ── Open SAS modal ─────────────────────────────────────────

const _SAS_PERM_LABELS = {
  r: "Read", a: "Add", c: "Create", w: "Write",
  d: "Delete", l: "List", t: "Tags", x: "Execute",
  f: "Find", m: "Move", e: "Permanent delete", i: "Set immutability",
  o: "Ownership", p: "Permissions",
};

function _showOpenSasModal() {
  const overlay   = _el("openSasModal");
  const input     = _el("openSasInput");
  const openBtn   = _el("openSasOpenBtn");
  const errEl     = _el("openSasError");
  const parsed    = _el("openSasParsed");
  const closeBtn  = _el("openSasModalClose");
  const cancelBtn = _el("openSasCancelBtn");

  // Reset state
  input.value = "";
  openBtn.disabled = true;
  errEl.classList.add("hidden");
  parsed.classList.add("hidden");
  overlay.classList.remove("hidden");
  input.focus();

  let _parsedSas = null;

  // Live-parse the URL as the user types/pastes (read-only — no global state mutation)
  input.oninput = () => {
    const raw = input.value.trim();
    errEl.classList.add("hidden");
    parsed.classList.add("hidden");
    openBtn.disabled = true;
    _parsedSas = null;

    if (!raw) return;

    try {
      const info = parseSasUrl(raw);
      // Show parsed info
      _el("openSasAccount").textContent   = info.accountName;
      _el("openSasContainer").textContent = info.containerName;
      _el("openSasPath").textContent      = info.blobPrefix || "(container root)";

      // Friendly permission labels
      const permLabels = info.permissions.split("").map(ch => _SAS_PERM_LABELS[ch] || ch).join(", ");
      _el("openSasPerms").textContent = permLabels || "(not specified)";

      _el("openSasStart").textContent  = info.start  ? info.start.toLocaleString()  : "(immediate)";
      _el("openSasExpiry").textContent = info.expiry ? info.expiry.toLocaleString() : "(not set)";

      const resourceLabels = { c: "Container", b: "Blob", d: "Directory" };
      _el("openSasResource").textContent = resourceLabels[info.signedResource] || info.signedResource || "(not specified)";

      parsed.classList.remove("hidden");
      openBtn.disabled = false;
      _parsedSas = { raw, info };
    } catch (e) {
      errEl.textContent = e.message;
      errEl.classList.remove("hidden");
    }
  };

  const close = () => {
    overlay.classList.add("hidden");
    input.oninput = null;
    openBtn.onclick = null;
    closeBtn.onclick = null;
    cancelBtn.onclick = null;
    if (_escHandler) { overlay.removeEventListener("keydown", _escHandler); }
  };

  openBtn.onclick = () => {
    if (!_parsedSas) return;
    close();
    // Activate SAS mode with the validated URL (only now mutating global state)
    activateSasMode(_parsedSas.raw);
    _bootSasMode(_parsedSas.info, _parsedSas.info);
  };

  closeBtn.onclick = close;
  cancelBtn.onclick = close;

  const _escHandler = (e) => { if (e.key === "Escape") close(); };
  overlay.addEventListener("keydown", _escHandler);
  overlay.addEventListener("click", (e) => { if (e.target === overlay) close(); }, { once: true });
}

/**
 * Boot the app in SAS-token browsing mode.
 * Similar to _bootApp but skips OAuth-dependent features.
 */
function _bootSasMode(info, state) {
  _el("signInPage").classList.add("hidden");
  _el("pickerPage").classList.add("hidden");
  _el("mainApp").classList.remove("hidden");
  _el("headerInfoBar").classList.remove("hidden");
  document.title = CONFIG.app.title;

  // Header info
  _el("userDisplayName").textContent = "";
  _el("userInfo").classList.add("hidden");
  _el("storageInfo").textContent =
    `${info.accountName} / ${info.containerName}`;

  // Hide OAuth-only controls
  _el("signOutBtn").classList.add("hidden");
  _el("changeStorageBtn").classList.add("hidden");

  // Show SAS badge with expiry info
  const badge = _el("permBadge");
  if (state.expiry) {
    const now = new Date();
    if (state.expiry < now) {
      badge.textContent = "🔑 SAS Expired";
      badge.className = "sas-badge sas-badge-expired";
    } else {
      badge.textContent = "🔑 SAS Active";
      badge.className = "sas-badge sas-badge-active";
    }
  } else {
    badge.textContent = "🔑 SAS";
    badge.className = "sas-badge sas-badge-active";
  }
  badge.classList.remove("hidden");

  // Add an "Open SAS" button in the header for switching SAS URLs & a "Disconnect" button
  let sasBtnHeader = document.getElementById("openSasHeaderBtn");
  if (!sasBtnHeader) {
    sasBtnHeader = document.createElement("button");
    sasBtnHeader.id = "openSasHeaderBtn";
    sasBtnHeader.className = "btn btn-secondary open-sas-header-btn";
    sasBtnHeader.title = "Open a different SAS URL";
    sasBtnHeader.innerHTML = "🔑 <span class='btn-label'>Open SAS</span>";
    _el("signOutBtn").parentNode.insertBefore(sasBtnHeader, _el("signOutBtn"));
  }
  sasBtnHeader.classList.remove("hidden");
  sasBtnHeader.onclick = () => _showOpenSasModal();

  let disconnectBtn = document.getElementById("sasDisconnectBtn");
  if (!disconnectBtn) {
    disconnectBtn = document.createElement("button");
    disconnectBtn.id = "sasDisconnectBtn";
    disconnectBtn.className = "btn btn-secondary";
    disconnectBtn.title = "Disconnect SAS and return to sign-in";
    disconnectBtn.innerHTML = "⏏ <span class='btn-label'>Disconnect</span>";
    sasBtnHeader.parentNode.insertBefore(disconnectBtn, sasBtnHeader.nextSibling);
  }
  disconnectBtn.classList.remove("hidden");
  disconnectBtn.onclick = () => {
    deactivateSasMode();
    // Reset SAS-specific UI
    badge.className = "perm-badge hidden";
    sasBtnHeader.classList.add("hidden");
    disconnectBtn.classList.add("hidden");

    // If there is still an active OAuth session, return to the main app / picker
    // rather than the sign-in page.
    const activeUser = getUser();
    if (activeUser) {
      _el("userInfo").classList.remove("hidden");
      _el("signOutBtn").classList.remove("hidden");
      const pickerAllowed = CONFIG.app.allowStoragePicker !== false;
      if (pickerAllowed) {
        _showPickerPage(activeUser);
      } else {
        _bootApp(activeUser);
      }
    } else {
      _el("mainApp").classList.add("hidden");
      _el("userInfo").classList.remove("hidden");
      _el("signOutBtn").classList.remove("hidden");
      _el("changeStorageBtn").classList.remove("hidden");
      _showSignInPage();
    }
  };

  // Derive permissions from SAS sp= parameter
  const perms = state.permissions || "";
  const canRead   = perms.includes("r");
  const canList   = perms.includes("l");
  const canWrite  = /[wca]/.test(perms);
  const canDelete = perms.includes("d");
  _canUpload = canWrite;

  // Show/hide upload controls based on write permission
  _el("uploadBtn").classList.toggle("hidden", !canWrite);
  _el("newBtn").classList.toggle("hidden", !canWrite);
  _el("auditBtn").classList.add("hidden"); // audit viewer not available in SAS mode
  _el("uploadPanel").classList.add("hidden");
  _el("uploadBtn").classList.remove("active");
  _uploadQueue   = [];
  _uploadCounter = 0;
  _el("uploadQueueList").innerHTML = "";
  _el("uploadQueue").classList.add("hidden");

  if (canWrite) {
    _el("newBtn").onclick         = _showNewModal;
    _el("uploadBtn").onclick      = _toggleUploadPanel;
    _el("pickFilesBtn").onclick   = () => _el("fileInput").click();
    _el("pickFolderBtn").onclick  = () => _el("folderInput").click();
    _el("fileInput").onchange     = (e) => { _queueFiles(e.target.files); e.target.value = ""; };
    _el("folderInput").onchange   = (e) => { _queueFiles(e.target.files); e.target.value = ""; };
    _el("clearCompletedBtn").onclick = _clearCompleted;
    _initDragDrop();
  }

  // Hide download-all if no read permission
  _el("downloadAllBtn").classList.toggle("hidden", !canRead);

  // Reset search / sort state
  _cachedFolders          = [];
  _cachedFiles            = [];
  _containerSearchResults = null;
  _el("searchInput").value = "";
  _el("searchClearBtn").classList.add("hidden");
  _el("searchBar").classList.add("hidden");
  _el("searchBtn").classList.remove("active");
  _el("searchBanner").classList.add("hidden");

  // Wire toolbar
  _el("refreshBtn").onclick     = () => _loadFiles(_currentPrefix);
  _el("reportBtn").onclick      = _exportReport;
  _el("downloadAllBtn").onclick = _downloadCurrentLevel;
  _el("infoBtn").onclick        = _showInfoModal;
  _el("upBtn").onclick          = _goUp;
  _el("footerYear").textContent = new Date().getFullYear();
  _el("listViewBtn").onclick    = () => _setViewMode("list");
  _el("gridViewBtn").onclick    = () => _setViewMode("grid");
  _el("searchBtn").onclick      = _toggleSearchBar;
  _el("searchInput").oninput    = _onSearchInput;
  _el("searchInput").onkeydown  = (e) => { if (e.key === "Escape") _toggleSearchBar(); };
  _el("searchClearBtn").onclick = _clearSearch;
  document.querySelectorAll("input[name='searchScope']").forEach((r) => { r.onchange = _onSearchInput; });

  // Initialize folder tree
  _initFolderTree();

  // Determine starting prefix
  _currentPrefix  = "";
  _listLoadFailed = false;
  _listingPromise = null;
  _selection.clear();

  const startPrefix = info.blobPrefix;

  // If the SAS has no list permission but has a specific blob path, skip
  // listing entirely and show a download prompt straight away.
  if (!canList && startPrefix && !startPrefix.endsWith("/")) {
    _showLoading(false);
    _showSasDownloadPanel(startPrefix);
    return;
  }

  // If the SAS is for a single blob (sr=b), navigate to its parent and auto-view/download
  if (state.signedResource === "b" && startPrefix && !startPrefix.endsWith("/")) {
    const parts = startPrefix.split("/");
    const blobName = parts.pop();
    const parentPrefix = parts.length ? parts.join("/") + "/" : "";
    _listingPromise = _loadFiles(parentPrefix);
    // Auto-show the blob after listing loads
    _listingPromise.then(() => {
      // Find the file in the cached list
      const file = _cachedFiles.find(f => f.name === startPrefix);
      if (file) _showViewModal(file);
    }).catch(() => {});
  } else {
    // Container or directory SAS — navigate to the prefix
    const prefix = startPrefix ? (startPrefix.endsWith("/") ? startPrefix : startPrefix + "/") : "";
    _listingPromise = _loadFiles(prefix);
  }

  _showLoading(false);
}

// ── Storage picker ───────────────────────────────────────────

/**
 * Show a simple download/view card in the content area when a SAS token
 * allows reading a specific blob but not listing the container.
 */
function _showSasDownloadPanel(blobPath) {
  const fileName = blobPath.split("/").pop();
  const icon     = getFileIcon(fileName);
  const viewable = _isViewable(fileName);

  _renderBreadcrumb("");
  _el("emptyState").classList.add("hidden");
  _el("fileContainer").innerHTML = "";
  _hideError();
  _el("networkErrorHelp").classList.add("hidden");

  const card = document.createElement("div");
  card.className = "sas-download-card";
  card.innerHTML = `
    <div class="sas-download-icon">${_esc(icon)}</div>
    <div class="sas-download-name">${_esc(fileName)}</div>
    <p class="sas-download-hint">The SAS token grants read access to this file. What would you like to do?</p>
    <div class="sas-download-actions">
      ${viewable ? `<button class="btn btn-secondary sas-dl-view-btn">👁 View</button>` : ""}
      <button class="btn btn-primary sas-dl-download-btn">⬇ Download</button>
    </div>
    <p class="sas-download-error hidden"></p>`;

  _el("fileContainer").appendChild(card);

  // Wire Download
  card.querySelector(".sas-dl-download-btn").addEventListener("click", async () => {
    try {
      await downloadBlob(blobPath);
    } catch (err) {
      const errEl = card.querySelector(".sas-download-error");
      errEl.textContent = err.message;
      errEl.classList.remove("hidden");
    }
  });

  // Wire View (if applicable)
  if (viewable) {
    card.querySelector(".sas-dl-view-btn").addEventListener("click", () => {
      _showViewModal({ name: blobPath, displayName: fileName, size: 0 });
    });
  }
}

/**
 * Show the storage account / container picker page.
 * Discovers all accessible storage accounts via the ARM API and
 * renders them as collapsible cards. Clicking a container boots
 * the main app with that account+container.
 */
async function _showPickerPage(account) {
  _el("signInPage").classList.add("hidden");
  _el("mainApp").classList.add("hidden");
  _el("pickerPage").classList.remove("hidden");
  _el("headerInfoBar").classList.add("hidden");
  document.title = "Select Storage — " + (CONFIG.app.title || "Blob Browser");

  // Populate user info in picker sub-header info bar
  const displayName = account.name || account.username;
  const upn          = account.username;
  _el("pickerInfoBarName").textContent = upn ? `${displayName} (${upn})` : displayName;
  _el("pickerSignOutBtn").onclick = () => signOut();
  _el("pickerOriginHint").textContent = window.location.origin;

  // Show the Back button only when the user is already inside a session
  // (i.e. they came here via the Change button, not on initial sign-in)
  const hasActiveSession = !!(CONFIG.storage.accountName && CONFIG.storage.containerName);
  const backBtn = _el("pickerBackBtn");
  backBtn.classList.toggle("hidden", !hasActiveSession);
  backBtn.onclick = hasActiveSession ? () => {
    _el("pickerPage").classList.add("hidden");
    _el("mainApp").classList.remove("hidden");
    _el("headerInfoBar").classList.remove("hidden");
  } : null;

  const subList  = _el("pickerSubList");
  const errorEl  = _el("pickerError");
  const emptyEl  = _el("pickerEmpty");
  const searchEl = _el("pickerSearch");

  subList.innerHTML = "";
  errorEl.classList.add("hidden");
  emptyEl.classList.add("hidden");
  searchEl.value = "";

  // Show loading skeletons
  for (let i = 0; i < 3; i++) {
    const sk = document.createElement("div");
    sk.className = "picker-skeleton";
    subList.appendChild(sk);
  }

  try {
    const subscriptions = await listSubscriptions();

    if (subscriptions.length === 0) {
      subList.innerHTML = "";
      emptyEl.classList.remove("hidden");
      return;
    }

    // Fetch storage accounts for all subscriptions in parallel
    const subResults = await Promise.allSettled(
      subscriptions.map(s => listStorageAccounts(s.subscriptionId).then(accounts => ({ sub: s, accounts })))
    );

    subList.innerHTML = "";
    let totalAccounts = 0;

    // Collect all accounts across subscriptions
    const accountEntries = [];
    for (const result of subResults) {
      if (result.status === "rejected" || !result.value.accounts.length) continue;
      const { sub, accounts } = result.value;
      for (const acct of accounts) accountEntries.push({ sub, acct });
    }

    // Fetch containers AND CORS settings for all accounts in parallel
    const [containerFetches, corsFetches] = await Promise.all([
      Promise.allSettled(
        accountEntries.map(({ acct }) =>
          listContainers(acct.subscriptionId, acct.resourceGroup, acct.name)
            .then(containers => ({ name: acct.name, containers }))
        )
      ),
      Promise.allSettled(
        accountEntries.map(({ acct }) =>
          getBlobServiceCors(acct.subscriptionId, acct.resourceGroup, acct.name)
            .then(allowed => ({ name: acct.name, allowed }))
        )
      ),
    ]);

    const containerMap = new Map();
    containerFetches.forEach(r => {
      if (r.status === "fulfilled") containerMap.set(r.value.name, r.value.containers);
    });

    const corsMap = new Map();
    corsFetches.forEach(r => {
      if (r.status === "fulfilled") corsMap.set(r.value.name, r.value.allowed);
    });

    // Group by subscription, skipping accounts with zero containers or no CORS
    const subGroups = new Map();
    for (const { sub, acct } of accountEntries) {
      const containers = containerMap.get(acct.name);
      if (containers !== undefined && containers.length === 0) continue;
      // Hide accounts where CORS is definitively not configured for this origin
      const corsOk = corsMap.get(acct.name);
      if (corsOk === false) continue;
      if (!subGroups.has(sub.subscriptionId)) subGroups.set(sub.subscriptionId, { sub, entries: [] });
      subGroups.get(sub.subscriptionId).entries.push({ acct, containers: containers ?? null });
    }

    for (const { sub, entries } of subGroups.values()) {
      const group = document.createElement("div");
      group.className = "picker-sub-group";
      group.dataset.sub = sub.subscriptionId;

      const label = document.createElement("div");
      label.className = "picker-sub-label";
      label.textContent = sub.displayName;
      group.appendChild(label);

      for (const { acct, containers } of entries) {
        group.appendChild(_makeAccountCard(acct, account, containers));
        totalAccounts++;
      }

      subList.appendChild(group);
    }

    if (totalAccounts === 0) {
      emptyEl.classList.remove("hidden");
      return;
    }

    // Expand account that matches current selection (if any)
    const sel = _getStorageSelection();
    if (sel) {
      const matching = subList.querySelector(`[data-account="${sel.accountName}"]`);
      if (matching) matching.classList.add("expanded");
    }

    // Wire search filter
    searchEl.oninput = () => _filterPicker(searchEl.value.trim().toLowerCase());

  } catch (err) {
    console.error("[picker] Failed:", err);
    subList.innerHTML = "";
    errorEl.textContent = `Failed to load storage accounts: ${err.message}`;
    errorEl.classList.remove("hidden");
    // Still allow manually picking if config has accountName set
    if (CONFIG.storage.accountName && CONFIG.storage.containerName) {
      const resumeBtn = document.createElement("button");
      resumeBtn.className = "btn btn-primary-sm";
      resumeBtn.style.marginTop = "12px";
      resumeBtn.textContent = `Continue with ${CONFIG.storage.accountName} / ${CONFIG.storage.containerName}`;
      resumeBtn.onclick = () => _bootApp(account);
      errorEl.appendChild(document.createElement("br"));
      errorEl.appendChild(resumeBtn);
    }
  }
}

function _makeAccountCard(acct, userAccount, prefetchedContainers = null) {
  const card = document.createElement("div");
  card.className = "picker-account";
  card.dataset.account = acct.name;

  const sel = _getStorageSelection();
  const isCurrentAccount = sel && sel.accountName === acct.name;

  // Header row (click to expand/collapse)
  const header = document.createElement("div");
  header.className = "picker-account-header";
  header.innerHTML = `
    <div class="picker-account-left">
      <span class="picker-account-icon">🗄️</span>
      <div>
        <div class="picker-account-name">${_esc(acct.name)}</div>
        <div class="picker-account-meta">${_esc(acct.location)} · ${_esc(acct.resourceGroup)}</div>
      </div>
    </div>
    <span class="picker-account-chevron">▶</span>`;
  header.addEventListener("click", () => {
    const isExpanded = card.classList.contains("expanded");
    if (!isExpanded && !card.dataset.loaded) {
      _loadContainersIntoCard(card, acct, userAccount, prefetchedContainers);
    }
    card.classList.toggle("expanded", !isExpanded);
  });
  card.appendChild(header);

  // Container list (initially hidden)
  const containerList = document.createElement("div");
  containerList.className = "picker-containers";
  containerList.dataset.role = "containers";
  containerList.innerHTML = `<div style="padding:12px 16px 12px 48px;font-size:12px;color:var(--text-muted)">Loading containers…</div>`;
  card.appendChild(containerList);

  // Pre-load containers for the current account, or when data is already available
  if (isCurrentAccount || prefetchedContainers !== null) {
    _loadContainersIntoCard(card, acct, userAccount, prefetchedContainers);
  }

  return card;
}

async function _loadContainersIntoCard(card, acct, userAccount, prefetchedContainers = null) {
  if (card.dataset.loaded) return;
  card.dataset.loaded = "1";

  const containerList = card.querySelector("[data-role='containers']");
  try {
    const containers = prefetchedContainers !== null
      ? prefetchedContainers
      : await listContainers(acct.subscriptionId, acct.resourceGroup, acct.name);
    containerList.innerHTML = "";

    if (containers.length === 0) {
      containerList.innerHTML = `<div style="padding:12px 16px 12px 48px;font-size:12px;color:var(--text-muted)">No containers found</div>`;
      return;
    }

    const sel = _getStorageSelection();

    for (const c of containers) {
      const item = document.createElement("div");
      item.className = "picker-container-item";
      if (sel && sel.accountName === acct.name && sel.containerName === c.name) {
        item.classList.add("current-selection");
      }
      item.innerHTML = `
        <div class="picker-container-left">
          <span>📦</span>
          <span class="picker-container-name">${_esc(c.name)}</span>
          ${c.publicAccess !== "None" ? `<span class="picker-container-badge">${_esc(c.publicAccess)}</span>` : ""}
        </div>
        <span class="picker-container-action">${sel && sel.accountName === acct.name && sel.containerName === c.name ? "Current" : "Open →"}</span>`;

      item.addEventListener("click", () => {
        // Apply selection
        _setStorageSelection(acct.name, c.name);
        _bootApp(userAccount);
      });
      containerList.appendChild(item);
    }
  } catch (err) {
    containerList.innerHTML = `<div style="padding:12px 16px 12px 48px;font-size:12px;color:var(--error)">Error loading containers: ${_esc(err.message)}</div>`;
  }
}

function _filterPicker(query) {
  const groups = _el("pickerSubList").querySelectorAll(".picker-sub-group");
  groups.forEach(group => {
    let groupVisible = false;
    group.querySelectorAll(".picker-account").forEach(card => {
      const accountName = (card.dataset.account || "").toLowerCase();
      const rg          = (card.querySelector(".picker-account-meta")?.textContent || "").toLowerCase();
      const accountMatch = !query || accountName.includes(query) || rg.includes(query);

      if (accountMatch) {
        // Account name / RG matches — show card and all its containers
        card.style.display = "";
        card.querySelectorAll(".picker-container-item").forEach(item => {
          item.style.display = "";
        });
        groupVisible = true;
      } else {
        // Account name doesn't match — filter individual containers
        let anyContainerMatch = false;
        card.querySelectorAll(".picker-container-item").forEach(item => {
          const cName = (item.querySelector(".picker-container-name")?.textContent || "").toLowerCase();
          const matches = cName.includes(query);
          item.style.display = matches ? "" : "none";
          if (matches) anyContainerMatch = true;
        });
        card.style.display = anyContainerMatch ? "" : "none";
        if (anyContainerMatch) {
          // Auto-expand so matching containers are visible
          card.classList.add("expanded");
        }
        if (anyContainerMatch) groupVisible = true;
      }
    });
    group.style.display = groupVisible ? "" : "none";
  });
}

// ── Storage selection persistence ────────────────────────────

function _setStorageSelection(accountName, containerName) {
  CONFIG.storage.accountName  = accountName;
  CONFIG.storage.containerName = containerName;
  sessionStorage.setItem(_STORAGE_SELECTION_KEY, JSON.stringify({ accountName, containerName }));
}

function _getStorageSelection() {
  const raw = sessionStorage.getItem(_STORAGE_SELECTION_KEY);
  if (!raw) return null;
  try {
    const sel = JSON.parse(raw);
    if (typeof sel?.accountName !== "string" || typeof sel?.containerName !== "string") return null;
    return { accountName: sel.accountName, containerName: sel.containerName };
  } catch {
    return null;
  }
}

function _restoreStorageSelection() {
  const sel = _getStorageSelection();
  if (sel) {
    CONFIG.storage.accountName   = sel.accountName;
    CONFIG.storage.containerName = sel.containerName;
  }
}

function _bootApp(account) {
  _el("signInPage").classList.add("hidden");
  _el("pickerPage").classList.add("hidden");
  _el("mainApp").classList.remove("hidden");
  _el("headerInfoBar").classList.remove("hidden");
  document.title = CONFIG.app.title;

  // Header info
  const displayName = account.name || account.username;
  const upn          = account.username;
  _el("userDisplayName").textContent = upn ? `${displayName} (${upn})` : displayName;
  _el("storageInfo").textContent =
    `${CONFIG.storage.accountName} / ${CONFIG.storage.containerName}`;

  // Reset upload state when switching containers
  _canUpload = false;
  _el("uploadBtn").classList.add("hidden");
  _el("newBtn").classList.add("hidden");
  _el("auditBtn").classList.add("hidden");
  _el("uploadPanel").classList.add("hidden");
  _el("uploadBtn").classList.remove("active");
  _uploadQueue   = [];
  _uploadCounter = 0;
  _el("uploadQueueList").innerHTML = "";
  _el("uploadQueue").classList.add("hidden");

  // Reset search / sort state when switching containers
  _cachedFolders          = [];
  _cachedFiles            = [];
  _containerSearchResults = null;
  _el("searchInput").value = "";
  _el("searchClearBtn").classList.add("hidden");
  _el("searchBar").classList.add("hidden");
  _el("searchBtn").classList.remove("active");
  _el("searchBanner").classList.add("hidden");

  // Wire up toolbar buttons (replace to avoid double-binding on container switch)
  // Show/hide Change Storage button based on config
  const pickerAllowed = CONFIG.app.allowStoragePicker !== false;
  _el("changeStorageBtn").classList.toggle("hidden", !pickerAllowed);

  // Ensure OAuth controls are visible (may have been hidden by a prior SAS session)
  _el("signOutBtn").classList.remove("hidden");
  _el("userInfo").classList.remove("hidden");
  document.getElementById("openSasHeaderBtn")?.classList.add("hidden");
  document.getElementById("sasDisconnectBtn")?.classList.add("hidden");

  // "Open SAS URL" button — add once, re-use on container switch
  let openSasBtn = document.getElementById("openSasHeaderBtn");
  if (!openSasBtn) {
    openSasBtn = document.createElement("button");
    openSasBtn.id = "openSasHeaderBtn";
    openSasBtn.className = "btn btn-secondary open-sas-header-btn";
    openSasBtn.title = "Open a SAS URL without signing out";
    openSasBtn.innerHTML = "\uD83D\uDD11 <span class='btn-label'>Open SAS</span>";
    _el("signOutBtn").parentNode.insertBefore(openSasBtn, _el("signOutBtn"));
  }
  openSasBtn.classList.remove("hidden");
  openSasBtn.onclick = () => _showOpenSasModal();

  _el("signOutBtn").onclick     = () => signOut();
  _el("helpBtn").onclick        = _showHelpModal;
  _el("refreshBtn").onclick     = () => _loadFiles(_currentPrefix);
  _el("reportBtn").onclick      = _exportReport;
  _el("downloadAllBtn").onclick = _downloadCurrentLevel;
  _el("infoBtn").onclick        = _showInfoModal;
  _el("upBtn").onclick          = _goUp;
  _el("changeStorageBtn").onclick = pickerAllowed ? () => {
    const acc = getUser();
    _el("mainApp").classList.add("hidden");
    _showPickerPage(acc || account);
  } : null;

  _el("footerYear").textContent = new Date().getFullYear();

  _el("listViewBtn").onclick = () => _setViewMode("list");
  _el("gridViewBtn").onclick = () => _setViewMode("grid");

  // Search bar controls (safe to re-assign on container switch)
  _el("searchBtn").onclick      = _toggleSearchBar;
  _el("searchInput").oninput    = _onSearchInput;
  _el("searchInput").onkeydown  = (e) => { if (e.key === "Escape") _toggleSearchBar(); };
  _el("searchClearBtn").onclick = _clearSearch;
  document.querySelectorAll("input[name='searchScope']").forEach((r) => { r.onchange = _onSearchInput; });

  // Initialize folder tree sidebar
  _initFolderTree();

  // Load from hash deep-link if present, otherwise root
  _currentPrefix  = "";
  _listLoadFailed = false;
  _listingPromise = null;
  _selection.clear();
  const deepLink = _parseAppHash();
  _listingPromise = _loadFiles(deepLink.path || "");

  // Probe write access in parallel with the initial file listing.
  // The Upload button appears only if the signed-in user has the
  // "Storage Blob Data Contributor" role (or higher) on the container.
  probeUploadPermission().then(async (hasPermission) => {
    // Wait for the listing to settle so we never race against a CORS/network failure
    // (e.g. "Same Origin Policy disallows reading the remote resource" means CORS
    // is not configured — in that case _listLoadFailed will be true after this await)
    if (_listingPromise) await _listingPromise.catch(() => {});
    if (_listLoadFailed) return;   // listing failed — hide badge, don't show misleading role
    const badge = _el("permBadge");
    badge.textContent = hasPermission ? "\u270f\ufe0f Contributor" : "\uD83D\uDCD6 Reader";
    badge.classList.remove("hidden");
    if (!hasPermission) return;
    _canUpload = true;
    _el("uploadBtn").classList.remove("hidden");
    _el("newBtn").classList.remove("hidden");
    _el("auditBtn").classList.remove("hidden");
    _el("auditBtn").onclick       = _showAuditModal;
    _el("newBtn").onclick         = _showNewModal;
    _el("uploadBtn").onclick      = _toggleUploadPanel;
    _el("pickFilesBtn").onclick   = () => _el("fileInput").click();
    _el("pickFolderBtn").onclick  = () => _el("folderInput").click();
    _el("fileInput").onchange     = (e) => { _queueFiles(e.target.files); e.target.value = ""; };
    _el("folderInput").onchange   = (e) => { _queueFiles(e.target.files); e.target.value = ""; };
    _el("clearCompletedBtn").onclick = _clearCompleted;
    _initDragDrop();
  }).catch(() => {
    // probeUploadPermission never actually throws, but guard defensively — keep badge hidden
  });
}

// ── File loading ─────────────────────────────────────────────

function _goUp() {
  if (!_currentPrefix) return;
  const parts = _currentPrefix.split("/").filter(Boolean);
  parts.pop();
  _loadFiles(parts.length ? parts.join("/") + "/" : "");
}

async function _loadFiles(prefix) {
  _listLoadFailed = false;
  _currentPrefix = prefix;
  history.replaceState(null, "", _buildAppHash(prefix || ""));
  _selection.clear();
  _updateSelectionBar();
  _el("upBtn").classList.toggle("hidden", !prefix);
  _showLoading(true);
  _hideError();

  // Navigation always leaves container-wide search mode
  _containerSearchResults = null;
  _el("searchBanner").classList.add("hidden");
  // Clear search text when navigating to avoid stale cross-folder filter
  if (_el("searchInput").value) {
    _el("searchInput").value = "";
    _el("searchClearBtn").classList.add("hidden");
  }

  try {
    const { folders, files } = await listBlobsAtPrefix(prefix);
    _renderBreadcrumb(prefix);
    _renderFiles(folders, files);
    _syncTreeToPrefix(prefix).catch(() => {});
  } catch (err) {
    console.error("[app] Load error:", err);

    // In SAS mode a 403 on listing most likely means the token has no 'l'
    // (list) permission but may still allow reading a specific blob.
    // Show a download prompt instead of a raw error message.
    if (isSasMode() && /403|Access denied/i.test(err.message)) {
      const state = getSasState();
      const blobPath = state.blobPrefix || prefix.replace(/\/$/, "");
      if (blobPath) {
        _showLoading(false);
        _showSasDownloadPanel(blobPath);
        return;
      }
    }

    _showError(`Failed to load files: ${err.message}`);
    _renderFiles([], []);
    _listLoadFailed = true;
    _el("permBadge").classList.add("hidden");  // hide badge — can't confirm any access level
    // A TypeError means fetch() itself failed — CORS not configured or network access blocked
    if (err instanceof TypeError || /networkError|network error|failed to fetch/i.test(err.message)) {
      _el("networkErrorOrigin").textContent = window.location.origin;
      _el("networkErrorHelp").classList.remove("hidden");
    }
  } finally {
    _showLoading(false);
  }
}

// ── Breadcrumb ───────────────────────────────────────────────

function _renderBreadcrumb(prefix) {
  const ol = _el("breadcrumb");
  ol.innerHTML = "";

  ol.appendChild(_makeCrumb("\uD83D\uDCE6 " + CONFIG.storage.containerName, "", !prefix));

  if (prefix) {
    const parts = prefix.split("/").filter(Boolean);
    let accumulated = "";
    parts.forEach((part, i) => {
      accumulated += part + "/";
      ol.appendChild(_makeCrumb(part, accumulated, i === parts.length - 1));
    });
  }
}

function _makeCrumb(label, prefix, isActive) {
  const li = document.createElement("li");
  li.className = "breadcrumb-item" + (isActive ? " active" : "");

  if (isActive) {
    li.textContent = label;
  } else {
    const a   = document.createElement("a");
    a.href    = "#";
    a.textContent = label;
    a.addEventListener("click", (e) => { e.preventDefault(); _loadFiles(prefix); });
    li.appendChild(a);
  }
  return li;
}

// ── Rendering dispatcher ─────────────────────────────────────

function _renderFiles(folders, files) {
  _cachedFolders = folders;
  _cachedFiles   = files;
  _applyAndRender();
}

// ── Sort helpers ─────────────────────────────────────────────

/** Build a sortable <th> HTML string for the list-view header. */
function _makeSortTh(label, key) {
  const active = _sortKey === key;
  const arrow  = active
    ? ` <span class="sort-arrow">${_sortDir === "asc" ? "\u25b2" : "\u25bc"}</span>`
    : "";
  return `<th class="th-sortable${active ? " th-sort-active" : ""}" data-sort="${key}">${label}${arrow}</th>`;
}

/** Toggle or switch the active sort column, then re-render. */
function _setSort(key) {
  if (_sortKey === key) {
    _sortDir = _sortDir === "asc" ? "desc" : "asc";
  } else {
    _sortKey = key;
    _sortDir = "asc";
  }
  _applyAndRender();
}

/** Sort folders and files according to the current sort state. */
function _sortItems(folders, files) {
  const dir    = _sortDir === "asc" ? 1 : -1;
  const byName = (a, b) =>
    dir * a.displayName.localeCompare(b.displayName, undefined, { sensitivity: "base" });

  const sortedFolders = [...folders].sort((a, b) => {
    if (_sortKey === "name") return byName(a, b);
    return 0; // folders have no size/date — keep original order for those columns
  });

  const sortedFiles = [...files].sort((a, b) => {
    switch (_sortKey) {
      case "name":     return byName(a, b);
      case "size":     return dir * (a.size - b.size);
      case "modified": return dir * (new Date(a.lastModified) - new Date(b.lastModified));
      case "created":  return dir * (new Date(a.createdOn)    - new Date(b.createdOn));
      default:         return 0;
    }
  });

  return { folders: sortedFolders, files: sortedFiles };
}

/**
 * Central render function — applies active sort + folder-scope name filter
 * (or shows container search results), then delegates to list/grid renderers.
 */
function _applyAndRender() {
  const container  = _el("fileContainer");
  const emptyState = _el("emptyState");
  const countLabel = _el("itemCount");
  const banner     = _el("searchBanner");
  container.innerHTML = "";

  // ── Container-wide search results ──────────────────────────
  if (_containerSearchResults !== null) {
    const term = (_el("searchInput")?.value || "").trim();
    const { folders: sortedFolders, files: sortedFiles } = _sortItems(
      _containerSearchResults.folders.filter((f) => !f.name.startsWith(_AUDIT_FOLDER + "/")),
      _containerSearchResults.files.filter((f) => f.displayName !== ".keep" && !/^\.upload-probe-\d+$/.test(f.displayName) && !f.name.startsWith(_AUDIT_FOLDER + "/"))
    );
    const totalCount = sortedFolders.length + sortedFiles.length;
    banner.classList.remove("hidden");
    banner.innerHTML =
      `\uD83D\uDD0D <strong>${totalCount}</strong> result${totalCount !== 1 ? "s" : ""} `
      + `for \u201c<strong>${_esc(term)}</strong>\u201d across the entire container`;
    countLabel.textContent = `${totalCount} result${totalCount !== 1 ? "s" : ""}`;

    if (totalCount === 0) {
      emptyState.classList.remove("hidden");
      emptyState.querySelector("p").textContent = `No results found matching \u201c${term}\u201d`;
      return;
    }
    emptyState.classList.add("hidden");
    if (_viewMode === "list") {
      _renderSearchResultsListView(container, sortedFolders, sortedFiles);
    } else {
      _renderGridView(container, sortedFolders, sortedFiles);
    }
    return;
  }

  // ── Normal folder view ──────────────────────────────────────
  banner.classList.add("hidden");
  const term = (_el("searchInput")?.value || "").trim().toLowerCase();
  let folders = _cachedFolders;
  let files   = _cachedFiles;

  // Hide internal placeholder, probe blobs, and audit folder
  folders = folders.filter((f) => f.displayName !== _AUDIT_FOLDER);
  files = files.filter((f) => f.displayName !== ".keep" && !/^\.upload-probe-\d+$/.test(f.displayName));

  if (term) {
    folders = folders.filter((f) => f.displayName.toLowerCase().includes(term));
    files   = files.filter((f)   => f.displayName.toLowerCase().includes(term));
  }

  const { folders: sortedFolders, files: sortedFiles } = _sortItems(folders, files);
  const total = sortedFolders.length + sortedFiles.length;

  if (total === 0) {
    emptyState.classList.remove("hidden");
    emptyState.querySelector("p").textContent =
      term ? `No items match \u201c${term}\u201d` : "This folder is empty";
    countLabel.textContent = term ? `No results for \u201c${term}\u201d` : "Empty folder";
    return;
  }

  emptyState.classList.add("hidden");
  const suffix = term ? ` matching \u201c${term}\u201d` : "";
  countLabel.textContent =
    `${sortedFolders.length} folder${sortedFolders.length !== 1 ? "s" : ""},  `
  + `${sortedFiles.length} file${sortedFiles.length !== 1 ? "s" : ""}`
  + suffix;

  if (_viewMode === "list") {
    _renderListView(container, sortedFolders, sortedFiles);
  } else {
    _renderGridView(container, sortedFolders, sortedFiles);
  }
}

// ── Search handlers ───────────────────────────────────────────

function _toggleSearchBar() {
  const nowHidden = _el("searchBar").classList.toggle("hidden");
  _el("searchBtn").classList.toggle("active", !nowHidden);
  if (nowHidden) {
    _clearSearch();
  } else {
    _el("searchInput").focus();
  }
}

function _clearSearch() {
  clearTimeout(_searchDebounceTimer);
  _el("searchInput").value = "";
  _el("searchClearBtn").classList.add("hidden");
  _containerSearchResults = null;
  _el("searchBanner").classList.add("hidden");
  _applyAndRender();
}

function _onSearchInput() {
  const term  = (_el("searchInput")?.value || "").trim();
  const scope = document.querySelector("input[name='searchScope']:checked")?.value || "folder";
  _el("searchClearBtn").classList.toggle("hidden", !term);
  clearTimeout(_searchDebounceTimer);

  if (!term) {
    _containerSearchResults = null;
    _el("searchBanner").classList.add("hidden");
    _applyAndRender();
    return;
  }

  if (scope === "folder") {
    _containerSearchResults = null;
    _el("searchBanner").classList.add("hidden");
    _applyAndRender();
  } else {
    // "Everything" search: debounce 500 ms before hitting the API
    _searchDebounceTimer = setTimeout(() => _doContainerSearch(term), 500);
  }
}

async function _doContainerSearch(term) {
  _showLoading(true);
  try {
    _containerSearchResults = await listAllBlobs(term);
    _applyAndRender();
  } catch (err) {
    _showError(`Search failed: ${err.message}`);
  } finally {
    _showLoading(false);
  }
}

// ── Shared action-button helpers ──────────────────────────────
// These helpers centralise the repeated HTML generation and event wiring
// that was previously duplicated across _makeFolderRow, _makeFileRow,
// _makeSearchFolderRow, _makeSearchResultRow, _makeFolderCard and _makeFileCard.

/**
 * Return the inner HTML for action buttons appropriate for a folder or file.
 * @param {"folder"|"file"} type
 * @param {string}  displayName  Used to test viewability for files
 * @param {object}  opts
 * @param {boolean} opts.compact  true → icon-only labels (grid cards)
 * @param {string}  opts.extraHtml  Additional button HTML prepended to the list
 */
function _actionButtonsHtml(type, displayName, { compact = false, extraHtml = "" } = {}) {
  const lbl = (icon, text) => compact ? icon : `${icon} ${text}`;
  const isFolder = type === "folder";
  return `${extraHtml}
      <button class="btn-action btn-props" title="Properties">${lbl("ℹ️", "Info")}</button>
      <button class="btn-action btn-copy-url" title="Copy URL">${lbl("🔗", "Copy URL")}</button>
      ${_canSas() ? `<button class="btn-action btn-sas" title="Generate SAS">${lbl("🔑", "SAS")}</button>` : ""}
      ${!isFolder && _isViewable(displayName) ? `<button class="btn-action btn-view" title="View file">${lbl("👁", "View")}</button>` : ""}
      ${!isFolder && _isViewable(displayName) && !_isImageFile(displayName) && _canEditItems() ? `<button class="btn-action btn-edit" title="Edit file">${lbl("📝", "Edit")}</button>` : ""}
      ${isFolder && CONFIG.app.allowDownload ? `<button class="btn-action btn-dl-folder" title="Download as ZIP">${lbl("&#x2B07;", "Download")}</button>` : ""}
      ${!isFolder && CONFIG.app.allowDownload ? `<button class="btn-action btn-dl" title="Download">${lbl("&#x2B07;", "Download")}</button>` : ""}
      ${_canCopyItems() ? `<button class="btn-action btn-clone" title="Copy to…">${lbl("📋", "Copy")}</button>` : ""}
      ${_canMoveItems() ? `<button class="btn-action btn-move" title="Move to…">${lbl("📦", "Move")}</button>` : ""}
      ${_canRenameItems() ? `<button class="btn-action btn-rename" title="Rename">${lbl("✏️", "Rename")}</button>` : ""}
      ${_canDeleteItems() ? `<button class="btn-action btn-action-danger btn-delete" title="Delete">${lbl("🗑️", "Delete")}</button>` : ""}`;
}

/**
 * Wire action-button event listeners on a row or card element.
 * Handles both folder and file items, and optionally calls stopPropagation
 * (needed for cards where clicks on buttons should not trigger the card click).
 *
 * @param {Element}          el    The row (<tr>) or card (<div>) element
 * @param {object}           item  Folder or file object
 * @param {"folder"|"file"}  type
 * @param {object}  opts
 * @param {boolean} opts.stopPropagation  true for card elements
 * @param {string}  opts.displayName      Override for viewability check (search results use derived name)
 */
function _wireItemActions(el, item, type, { stopPropagation = false, displayName } = {}) {
  const isFolder = type === "folder";
  const viewName = displayName ?? item.displayName;

  const on = (sel, fn) => {
    const btn = el.querySelector(sel);
    if (!btn) return;
    btn.addEventListener("click", (e) => {
      if (stopPropagation) e.stopPropagation();
      fn(e);
    });
  };

  on(".btn-props",    () => isFolder ? _showFolderProperties(item) : _showProperties(item));
  on(".btn-copy-url", (e) => _showCopyMenu(e.currentTarget, item, type));
  on(".btn-sas",      () => _showSasModal(item, isFolder));
  if (isFolder) {
    on(".btn-dl-folder", () => _handleFolderDownload(item));
  } else {
    if (_isViewable(viewName))                                        on(".btn-view", () => _showViewModal(item));
    if (_isViewable(viewName) && !_isImageFile(viewName) && _canEditItems()) on(".btn-edit", () => _showEditModal(item));
    on(".btn-dl", () => _handleDownload(item));
  }
  on(".btn-clone",  () => _showCopyModal(item, type));
  on(".btn-move",   () => _showMoveModal(item, type));
  on(".btn-rename", () => _showRenameModal(item, type));
  on(".btn-delete", () => _showDeleteModal(item, type));
}

/** Wire a row-level checkbox (table rows). */
function _wireRowCheckbox(tr, name) {
  tr.querySelector(".row-chk").addEventListener("change", (e) => {
    if (e.target.checked) _selection.add(name); else _selection.delete(name);
    tr.classList.toggle("row-selected", e.target.checked);
    _updateSelectionBar();
  });
}

/** Wire a card-level checkbox (grid cards — includes stopPropagation). */
function _wireCardCheckbox(div, name) {
  div.querySelector(".card-chk").addEventListener("change", (e) => {
    e.stopPropagation();
    if (e.target.checked) _selection.add(name); else _selection.delete(name);
    div.classList.toggle("card-selected", e.target.checked);
    _updateSelectionBar();
  });
}

// ── List view ────────────────────────────────────────────────

function _renderListView(container, folders, files) {
  const table = document.createElement("table");
  table.className = "file-table";

  const thead = document.createElement("thead");
  thead.innerHTML = `
    <tr>
      <th class="col-check"><input type="checkbox" id="selectAllChk" title="Select all" /></th>
      ${_makeSortTh("Name",     "name")}
      ${_makeSortTh("Size",     "size")}
      ${_makeSortTh("Modified", "modified")}
      ${_makeSortTh("Created",  "created")}
      <th></th>
    </tr>`;
  table.appendChild(thead);

  // Wire sort-column click handlers
  thead.querySelectorAll("th.th-sortable").forEach((th) =>
    th.addEventListener("click", () => _setSort(th.dataset.sort))
  );

  const tbody = document.createElement("tbody");
  if (_currentPrefix) tbody.appendChild(_makeUpRow());
  folders.forEach((f) => tbody.appendChild(_makeFolderRow(f)));
  files.forEach((f)   => tbody.appendChild(_makeFileRow(f)));
  table.appendChild(tbody);
  container.appendChild(table);

  // Wire select-all
  const selectAllChk = table.querySelector("#selectAllChk");
  selectAllChk.addEventListener("change", () => {
    table.querySelectorAll(".row-chk").forEach((chk) => {
      chk.checked = selectAllChk.checked;
      const name = chk.dataset.name;
      if (selectAllChk.checked) _selection.add(name); else _selection.delete(name);
      chk.closest("tr").classList.toggle("row-selected", selectAllChk.checked);
    });
    _updateSelectionBar();
  });
}

function _parentPrefix() {
  const parts = _currentPrefix.split("/").filter(Boolean);
  parts.pop();
  return parts.length ? parts.join("/") + "/" : "";
}

function _makeUpRow() {
  const tr = document.createElement("tr");
  tr.className = "file-row up-row";
  tr.innerHTML = `
    <td class="col-check"></td>
    <td colspan="4">
      <div class="file-name">
        <span class="file-icon">&#x21A9;</span>
        <a href="#" class="file-link up-link">..</a>
      </div>
    </td>
    <td></td>`;
  tr.querySelector(".up-link").addEventListener("click", (e) => {
    e.preventDefault();
    _loadFiles(_parentPrefix());
  });
  return tr;
}

function _makeUpCard() {
  const div = document.createElement("div");
  div.className = "file-card up-card";
  div.setAttribute("role", "button");
  div.setAttribute("tabindex", "0");
  div.innerHTML = `
    <div class="card-icon">&#x21A9;</div>
    <div class="card-name">..</div>
    <div class="card-meta">Parent folder</div>`;
  const go = () => _loadFiles(_parentPrefix());
  div.addEventListener("click", go);
  div.addEventListener("keydown", (e) => {
    if (e.key === "Enter" || e.key === " ") { e.preventDefault(); go(); }
  });
  return div;
}

function _makeFolderRow(folder) {
  const tr = document.createElement("tr");
  tr.className = "file-row folder-row";
  if (_selection.has(folder.name)) tr.classList.add("row-selected");
  tr.innerHTML = `
    <td class="col-check"><input type="checkbox" class="row-chk" data-name="${_esc(folder.name)}" ${_selection.has(folder.name) ? "checked" : ""} /></td>
    <td>
      <div class="file-name">
        <span class="file-icon">📁</span>
        <a href="#" class="file-link">${_esc(folder.displayName)}</a>
      </div>
    </td>
    <td class="file-size">—</td>
    <td class="file-date">—</td>
    <td class="file-date">—</td>
    <td class="file-actions">
      <div class="row-actions">
        ${_actionButtonsHtml("folder", folder.displayName)}
      </div>
    </td>`;
  tr.querySelector(".file-link").addEventListener("click", (e) => {
    e.preventDefault();
    _loadFiles(folder.name);
  });
  _wireRowCheckbox(tr, folder.name);
  _wireItemActions(tr, folder, "folder");
  return tr;
}

/** Render a small metadata subtitle for a file row (uploaded-by / last-edited-by). */
function _fileMetaSubtitle(file) {
  const parts = [];
  if (file.uploadedByUpn)   parts.push(`⬆&thinsp;${_esc(file.uploadedByUpn)}`);
  if (file.lastEditedByUpn) parts.push(`✏&thinsp;${_esc(file.lastEditedByUpn)}`);
  if (!parts.length) return "";
  return `<span class="file-meta-info">${parts.join(" &middot; ")}</span>`;
}

function _makeFileRow(file) {
  const tr = document.createElement("tr");
  tr.className = "file-row";
  if (_selection.has(file.name)) tr.classList.add("row-selected");
  tr.innerHTML = `
    <td class="col-check"><input type="checkbox" class="row-chk" data-name="${_esc(file.name)}" ${_selection.has(file.name) ? "checked" : ""} /></td>
    <td>
      <div class="file-name">
        <span class="file-icon">${_esc(getFileIcon(file.displayName))}</span>
        <div class="file-name-group">
          <span>${_esc(file.displayName)}</span>
          ${_fileMetaSubtitle(file)}
        </div>
      </div>
    </td>
    <td class="file-size">${formatFileSize(file.size)}</td>
    <td class="file-date">${formatDate(file.lastModified)}</td>
    <td class="file-date">${formatDate(file.createdOn)}</td>
    <td class="file-actions">
      <div class="row-actions">
        ${_actionButtonsHtml("file", file.displayName)}
      </div>
    </td>`;
  _wireRowCheckbox(tr, file.name);
  _wireItemActions(tr, file, "file");
  return tr;
}

// ── Search-results list view ──────────────────────────────────

/** List view for whole-container search results (flat rows, full blob paths). */
function _renderSearchResultsListView(container, folders, files) {
  const table = document.createElement("table");
  table.className = "file-table";

  const thead = document.createElement("thead");
  thead.innerHTML = `
    <tr>
      <th class="col-check"><input type="checkbox" id="selectAllChk" title="Select all" /></th>
      ${_makeSortTh("Path",     "name")}
      ${_makeSortTh("Size",     "size")}
      ${_makeSortTh("Modified", "modified")}
      ${_makeSortTh("Created",  "created")}
      <th></th>
    </tr>`;
  table.appendChild(thead);
  thead.querySelectorAll("th.th-sortable").forEach((th) =>
    th.addEventListener("click", () => _setSort(th.dataset.sort))
  );

  const tbody = document.createElement("tbody");
  folders.forEach((f) => tbody.appendChild(_makeSearchFolderRow(f)));
  files.forEach((f) => tbody.appendChild(_makeSearchResultRow(f)));
  table.appendChild(tbody);
  container.appendChild(table);

  const selectAllChk = table.querySelector("#selectAllChk");
  selectAllChk.addEventListener("change", () => {
    table.querySelectorAll(".row-chk").forEach((chk) => {
      chk.checked = selectAllChk.checked;
      if (selectAllChk.checked) _selection.add(chk.dataset.name);
      else _selection.delete(chk.dataset.name);
      chk.closest("tr").classList.toggle("row-selected", selectAllChk.checked);
    });
    _updateSelectionBar();
  });
}

/** Build a table row for a folder found in the whole-container search. */
function _makeSearchFolderRow(folder) {
  const parentPrefix = folder.name.slice(0, folder.name.slice(0, -1).lastIndexOf("/") + 1);

  const tr = document.createElement("tr");
  tr.className = "file-row folder-row search-result-row";
  if (_selection.has(folder.name)) tr.classList.add("row-selected");
  tr.innerHTML = `
    <td class="col-check"><input type="checkbox" class="row-chk" data-name="${_esc(folder.name)}" ${_selection.has(folder.name) ? "checked" : ""} /></td>
    <td>
      <div class="file-name">
        <span class="file-icon">📁</span>
        <div class="search-result-name">
          <a href="#" class="file-link search-result-file">${_esc(folder.displayName)}</a>
          ${parentPrefix ? `<span class="search-result-path">${_esc(parentPrefix)}</span>` : ""}
        </div>
      </div>
    </td>
    <td class="file-size">—</td>
    <td class="file-date">—</td>
    <td class="file-date">—</td>
    <td class="file-actions">
      <div class="row-actions">
        ${_actionButtonsHtml("folder", folder.displayName)}
      </div>
    </td>`;

  tr.querySelector(".file-link").addEventListener("click", (e) => {
    e.preventDefault();
    _clearSearch();
    _loadFiles(folder.name);
  });
  _wireRowCheckbox(tr, folder.name);
  _wireItemActions(tr, folder, "folder");
  return tr;
}

/** Build a table row for a single whole-container search result. */
function _makeSearchResultRow(file) {
  const lastSlash    = file.name.lastIndexOf("/");
  const parentPrefix = lastSlash >= 0 ? file.name.slice(0, lastSlash + 1) : "";
  const fileName     = lastSlash >= 0 ? file.name.slice(lastSlash + 1) : file.name;

  const tr = document.createElement("tr");
  tr.className = "file-row search-result-row";
  if (_selection.has(file.name)) tr.classList.add("row-selected");
  tr.innerHTML = `
    <td class="col-check"><input type="checkbox" class="row-chk" data-name="${_esc(file.name)}" ${_selection.has(file.name) ? "checked" : ""} /></td>
    <td>
      <div class="file-name">
        <span class="file-icon">${_esc(getFileIcon(fileName))}</span>
        <div class="search-result-name">
          <span class="search-result-file">${_esc(fileName)}</span>
          ${parentPrefix ? `<span class="search-result-path">${_esc(parentPrefix)}</span>` : ""}
          ${_fileMetaSubtitle(file)}
        </div>
      </div>
    </td>
    <td class="file-size">${formatFileSize(file.size)}</td>
    <td class="file-date">${formatDate(file.lastModified)}</td>
    <td class="file-date">${formatDate(file.createdOn)}</td>
    <td class="file-actions">
      <div class="row-actions">
        ${_actionButtonsHtml("file", fileName, {
          extraHtml: parentPrefix ? `<button class="btn-action btn-goto-folder" title="Open containing folder">📂 Folder</button>` : "",
        })}
      </div>
    </td>`;

  _wireRowCheckbox(tr, file.name);
  if (parentPrefix) {
    tr.querySelector(".btn-goto-folder").addEventListener("click", () => {
      _clearSearch();
      _loadFiles(parentPrefix);
    });
  }
  _wireItemActions(tr, file, "file", { displayName: fileName });
  return tr;
}

// ── Grid view ────────────────────────────────────────────────

function _renderGridView(container, folders, files) {
  const grid = document.createElement("div");
  grid.className = "file-grid";
  if (_currentPrefix) grid.appendChild(_makeUpCard());
  folders.forEach((f) => grid.appendChild(_makeFolderCard(f)));
  files.forEach((f)   => grid.appendChild(_makeFileCard(f)));
  container.appendChild(grid);
}

function _makeFolderCard(folder) {
  const div = document.createElement("div");
  div.className = "file-card folder-card";
  if (_selection.has(folder.name)) div.classList.add("card-selected");
  div.setAttribute("role", "button");
  div.setAttribute("tabindex", "0");
  div.innerHTML = `
    <input type="checkbox" class="card-chk" data-name="${_esc(folder.name)}" ${_selection.has(folder.name) ? "checked" : ""} title="Select" />
    <div class="card-icon">📁</div>
    <div class="card-name" title="${_esc(folder.name)}">${_esc(folder.displayName)}</div>
    <div class="card-meta">Folder</div>
    <div class="card-actions">
      ${_actionButtonsHtml("folder", folder.displayName, { compact: true })}
    </div>`;
  div.addEventListener("click", (e) => {
    if (e.target.closest(".btn-action") || e.target.closest(".card-chk")) return;
    _loadFiles(folder.name);
  });
  div.addEventListener("keydown", (e) => {
    if (e.key === "Enter" || e.key === " ") { e.preventDefault(); _loadFiles(folder.name); }
  });
  _wireCardCheckbox(div, folder.name);
  _wireItemActions(div, folder, "folder", { stopPropagation: true });
  return div;
}

function _makeFileCard(file) {
  const div = document.createElement("div");
  div.className = "file-card";
  if (_selection.has(file.name)) div.classList.add("card-selected");
  div.innerHTML = `
    <input type="checkbox" class="card-chk" data-name="${_esc(file.name)}" ${_selection.has(file.name) ? "checked" : ""} title="Select" />
    <div class="card-icon">${_esc(getFileIcon(file.displayName))}</div>
    <div class="card-name" title="${_esc(file.name)}">${_esc(file.displayName)}</div>
    <div class="card-meta">${formatFileSize(file.size)}</div>
    <div class="card-meta card-date">${file.createdOn ? `Created ${formatDateShort(file.createdOn)}` : ""}</div>
    <div class="card-actions">
      ${_actionButtonsHtml("file", file.displayName, { compact: true })}
    </div>`;
  _wireCardCheckbox(div, file.name);
  _wireItemActions(div, file, "file", { stopPropagation: true });
  return div;
}

// ── File viewer ────────────────────────────────────────

const _VIEWABLE_EXTENSIONS = new Set([
  // Plain text
  "txt","log","md","markdown","rst","text","nfo",
  // Data / config
  "json","jsonc","yaml","yml","toml","xml","csv","tsv","ini","cfg","conf","env",
  // Web
  "html","htm","css","svg",
  // Code
  "js","ts","jsx","tsx","mjs","cjs",
  "py","rb","php","java","c","cpp","h","cs","go","rs","kt","swift",
  "sh","bash","zsh","ps1","bat","cmd","fish",
  "sql","graphql","gql",
  "gitignore","gitattributes","editorconfig","dockerfile",
]);

function _isViewable(filename) {
  if (_isImageFile(filename)) return true;
  const lower = filename.toLowerCase();
  // Match files with no extension that are known config files
  if (["dockerfile","makefile",".env",".gitignore",".gitattributes"].includes(lower)) return true;
  const ext = lower.split(".").pop();
  return _VIEWABLE_EXTENSIONS.has(ext);
}

const _MAX_VIEW_BYTES = 2 * 1024 * 1024; // 2 MB preview cap (text)

const _IMAGE_EXTENSIONS = new Set([
  "jpg","jpeg","png","gif","svg","webp","bmp","ico","avif",
]);

const _MAX_IMAGE_BYTES = 10 * 1024 * 1024; // 10 MB image preview cap

function _isImageFile(filename) {
  const lower = filename.toLowerCase();
  const dotIdx = lower.lastIndexOf(".");
  if (dotIdx < 0) return false;
  const ext = lower.slice(dotIdx + 1);
  return _IMAGE_EXTENSIONS.has(ext);
}

async function _showViewModal(file) {
  const modal = _el("viewModal");
  const body  = _el("viewModalBody");

  // Reset state — revoke any leftover blob URL from a previous preview
  const prevImg = body.querySelector(".view-image-preview");
  if (prevImg && prevImg.src.startsWith("blob:")) URL.revokeObjectURL(prevImg.src);
  body.innerHTML = `<div class="view-loading"><span class="view-spinner"></span> Loading…</div>`;
  _el("viewModalTitle").textContent = file.displayName;
  const ext = file.displayName.split(".").pop().toLowerCase();
  _el("viewLangBadge").textContent = ext.toUpperCase();
  _el("viewFileMeta").textContent  = "";
  _el("viewWrapBtn").classList.remove("active");
  modal.classList.remove("hidden");

  const isImage = _isImageFile(file.displayName);

  // Hide text-only controls for images
  _el("viewWrapBtn").classList.toggle("hidden", isImage);
  _el("viewCopyBtn").classList.toggle("hidden", isImage);

  let content = null;
  let wrapped  = false;

  const close = () => {
    modal.classList.add("hidden");
    // Revoke any blob URL to free memory
    const img = body.querySelector(".view-image-preview");
    if (img && img.src.startsWith("blob:")) URL.revokeObjectURL(img.src);
    body.innerHTML = "";
  };
  _el("viewModalClose").onclick = close;
  _el("viewCloseBtn").onclick   = close;
  modal.onclick = (e) => { if (e.target === modal) close(); };

  try {
    const { accountName, containerName } = CONFIG.storage;
    const authHeaders = await _storageAuthHeaders();
    const rawUrl = `https://${accountName}.blob.core.windows.net/${containerName}/${_encodePath(file.name)}`;
    const url    = _sasUrl(rawUrl);

    const res = await fetch(url, { headers: authHeaders });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);

    // Check size before reading body
    const cl = parseInt(res.headers.get("content-length") || "0", 10);
    const maxBytes = isImage ? _MAX_IMAGE_BYTES : _MAX_VIEW_BYTES;
    if (cl > maxBytes) {
      body.innerHTML = `<div class="view-too-large">
        <div style="font-size:32px">&#x26A0;&#xFE0F;</div>
        <p>File is too large to preview (${formatFileSize(cl)}).</p>
        ${CONFIG.app.allowDownload ? `<button id="viewDlInstead" class="btn btn-primary-sm">⬇ Download instead</button>` : ""}
      </div>`;
      if (CONFIG.app.allowDownload) _el("viewDlInstead").onclick = () => { close(); _handleDownload(file); };
      return;
    }

    if (isImage) {
      const blob = await res.blob();
      const blobUrl = URL.createObjectURL(blob);
      const container = document.createElement("div");
      container.className = "view-image-container";
      const img = document.createElement("img");
      img.className = "view-image-preview";
      img.alt = file.displayName;
      img.src = blobUrl;
      container.appendChild(img);
      body.innerHTML = "";
      body.appendChild(container);
      _el("viewFileMeta").textContent = formatFileSize(file.size || cl);
    } else {
      content = await res.text();
      const lineCount = content.split("\n").length;
      _el("viewFileMeta").textContent = `${lineCount.toLocaleString()} line${lineCount !== 1 ? "s" : ""} · ${formatFileSize(file.size)}`;

      _renderViewContent(body, content, wrapped);

      _el("viewWrapBtn").onclick = () => {
        wrapped = !wrapped;
        _el("viewWrapBtn").classList.toggle("active", wrapped);
        _renderViewContent(body, content, wrapped);
      };

      _el("viewCopyBtn").onclick = () => _copyToClipboard(content);
    }

    _el("viewDlBtn").classList.toggle("hidden", !CONFIG.app.allowDownload);
    _el("viewDlBtn").onclick = () => _handleDownload(file);

  } catch (err) {
    body.innerHTML = `<div class="view-too-large" style="color:var(--error)">Failed to load file: ${_esc(err.message)}</div>`;
  }
}

function _renderViewContent(body, content, wrapped) {
  // Use _esc() for robust HTML entity encoding (handles &, <, >, ", ')
  const safe = _esc(content);

  const lines    = safe.split("\n");
  const lineNums = lines.map((_, i) => i + 1).join("\n");

  body.innerHTML = `
    <div class="view-wrapper">
      <pre class="view-line-nums" aria-hidden="true">${lineNums}</pre>
      <pre class="view-pre${wrapped ? " view-pre-wrap" : ""}">${safe}</pre>
    </div>`;
}

// ── File editor ────────────────────────────────────────────

async function _showEditModal(file) {
  const modal    = _el("editModal");
  const textarea = _el("editContent");
  const saveBtn  = _el("editSaveBtn");
  const metaEl   = _el("editFileMeta");

  // Reset
  textarea.value = "";
  textarea.disabled = true;
  saveBtn.disabled = true;
  _el("editModalTitle").textContent = file.displayName;
  const ext = file.displayName.split(".").pop().toLowerCase();
  _el("editLangBadge").textContent = ext.toUpperCase();
  metaEl.textContent = "Loading…";
  modal.classList.remove("hidden");

  const close = () => { modal.classList.add("hidden"); textarea.value = ""; };
  _el("editModalClose").onclick = close;
  _el("editCancelBtn").onclick  = close;
  modal.onclick = (e) => { if (e.target === modal) close(); };

  try {
    const { accountName, containerName } = CONFIG.storage;
    const authHeaders = await _storageAuthHeaders();
    const rawUrl = `https://${accountName}.blob.core.windows.net/${containerName}/${_encodePath(file.name)}`;
    const url    = _sasUrl(rawUrl);
    const res    = await fetch(url, { headers: authHeaders });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const cl = parseInt(res.headers.get("content-length") || "0", 10);
    if (cl > _MAX_VIEW_BYTES) throw new Error(`File too large to edit (${formatFileSize(cl)})`);
    const text = await res.text();
    textarea.value = text;
    textarea.disabled = false;
    saveBtn.disabled = false;
    const lineCount = text.split("\n").length;
    metaEl.textContent = `${lineCount.toLocaleString()} line${lineCount !== 1 ? "s" : ""} · ${formatFileSize(file.size)}`;
  } catch (err) {
    metaEl.textContent = `Failed to load: ${err.message}`;
    return;
  }

  saveBtn.onclick = async () => {
    const newText = textarea.value;
    saveBtn.disabled = true;
    saveBtn.textContent = "Saving…";
    try {
      const mimeType = _guessMimeType(file.displayName);
      const newFile  = new File([newText], file.displayName, { type: mimeType });
      const user     = getUser();

      // Preserve existing metadata; stamp last_edited_by_*
      let meta = {};
      try { meta = await getBlobMetadata(file.name); } catch (e) { console.warn("[edit] Could not fetch metadata:", e.message); }
      const upn = user?.username || "";
      const oid = user?.oid      || "";
      if (upn) meta.last_edited_by_upn = upn;
      if (oid) meta.last_edited_by_oid = oid;

      await uploadBlob(file.name, newFile, null, meta);
      _audit("edit", file.name);
      close();
      _loadFiles(_currentPrefix);
      _showToast(`✅ "${file.displayName}" saved`);
    } catch (err) {
      metaEl.textContent = `Save failed: ${err.message}`;
      saveBtn.disabled = false;
      saveBtn.textContent = "💾 Save";
    }
  };
}
// ── View mode toggle ─────────────────────────────────────────

function _setViewMode(mode) {
  _viewMode = mode;
  _el("listViewBtn").classList.toggle("active", mode === "list");
  _el("gridViewBtn").classList.toggle("active", mode === "grid");
  _applyAndRender(); // re-render from cache — no network call, preserves search state
}

// ── CSV Report ────────────────────────────────────────────

/**
 * Open the report viewer modal. Shows a summary + table of all items
 * under the current prefix, with a download-CSV button.
 */
function _exportReport() {
  const modal   = _el("reportModal");
  const body    = _el("reportModalBody");
  const metaEl  = _el("reportMeta");
  const dlBtn   = _el("reportDownloadBtn");
  const genBtn  = _el("reportGenerateBtn");

  modal.classList.remove("hidden");
  body.innerHTML = `<div class="report-empty">Click <strong>Generate</strong> to scan <strong>${_esc(_currentPrefix || "/")}</strong>.</div>`;
  metaEl.textContent = "";
  dlBtn.disabled = true;

  // State
  let _allItems   = [];   // { type, displayName, name, size, contentType, lastModified, createdOn, etag, md5 }
  let _csvBlob    = null;
  let _csvName    = "";
  let _filterType = "";
  let _filterText = "";

  // Close
  const close = () => { modal.classList.add("hidden"); body.innerHTML = ""; _csvBlob = null; };
  _el("reportModalClose").onclick = close;
  _el("reportCloseBtn").onclick   = close;
  modal.onclick = (e) => { if (e.target === modal) close(); };

  // Download CSV
  dlBtn.onclick = () => {
    if (!_csvBlob) return;
    const url    = URL.createObjectURL(_csvBlob);
    const anchor = document.createElement("a");
    anchor.href     = url;
    anchor.download = _csvName;
    document.body.appendChild(anchor);
    anchor.click();
    document.body.removeChild(anchor);
    URL.revokeObjectURL(url);
    _showToast(`\u2705 Downloaded ${_csvName}`);
  };

  // Generate
  genBtn.onclick = generate;

  async function generate() {
    genBtn.disabled = true;
    genBtn.textContent = "Scanning\u2026";
    dlBtn.disabled = true;
    body.innerHTML = `<div class="report-loading"><span class="report-spinner"></span> Scanning folders\u2026</div>`;
    metaEl.textContent = "";
    _allItems = [];
    _csvBlob  = null;
    _filterType = "";
    _filterText = "";

    try {
      const { accountName, containerName } = CONFIG.storage;

      async function collect(prefix) {
        const { folders, files } = await listBlobsAtPrefix(prefix);
        for (const folder of folders) {
          _allItems.push({
            type:        "folder",
            displayName: folder.displayName,
            name:        folder.name,
            size:        0,
            contentType: "",
            lastModified: "",
            createdOn:   "",
            etag:        "",
            md5:         "",
          });
        }
        for (const file of files) {
          _allItems.push({
            type:        "file",
            displayName: file.displayName,
            name:        file.name,
            size:        file.size || 0,
            contentType: file.contentType || "",
            lastModified: file.lastModified || "",
            createdOn:   file.createdOn || "",
            etag:        file.etag || "",
            md5:         file.md5 || "",
          });
        }
        await Promise.all(folders.map(f => collect(f.name)));
      }

      await collect(_currentPrefix);

      // Build CSV blob for download
      const csvHeader = [
        "Type", "Name", "Full Path", "Path Length", "Blob URL",
        "Size (MB)", "Size (bytes)", "Content Type",
        "Last Modified", "Created On", "ETag", "MD5",
      ];
      const csvRows = [csvHeader];
      for (const item of _allItems) {
        const encoded = item.name.split("/").map(encodeURIComponent).join("/");
        const blobUrl = `https://${accountName}.blob.core.windows.net/${containerName}/${encoded}`;
        const sizeMB  = item.size > 0 ? (item.size / 1048576).toFixed(6) : "0";
        csvRows.push([
          item.type, item.displayName, item.name, String(item.name.length),
          blobUrl, sizeMB, String(item.size), item.contentType,
          item.lastModified, item.createdOn, item.etag, item.md5,
        ]);
      }
      const csvText = csvRows.map(r =>
        r.map(c => { const s = String(c ?? ""); return /[,"\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s; }).join(",")
      ).join("\r\n");

      const now  = new Date().toISOString().slice(0, 19).replace(/[T:]/g, "-");
      const loc  = _currentPrefix ? _currentPrefix.replace(/\/$/, "").split("/").pop() : containerName;
      _csvName = `report_${loc}_${now}.csv`;
      _csvBlob = new Blob(["\uFEFF" + csvText], { type: "text/csv;charset=utf-8;" });
      dlBtn.disabled = false;

      _renderReportTable();
    } catch (err) {
      console.error("[report]", err);
      body.innerHTML = `<div class="report-empty">\u26A0\uFE0F Report failed:<br><code>${_esc(err.message)}</code></div>`;
      metaEl.textContent = "";
    } finally {
      genBtn.disabled = false;
      genBtn.innerHTML = "\uD83D\uDCCA Generate";
    }
  }

  function _renderReportTable() {
    let items = _allItems;

    // Apply filters
    if (_filterType) items = items.filter(i => i.type === _filterType);
    if (_filterText) {
      const lower = _filterText.toLowerCase();
      items = items.filter(i =>
        i.displayName.toLowerCase().includes(lower) ||
        i.name.toLowerCase().includes(lower) ||
        i.contentType.toLowerCase().includes(lower)
      );
    }

    // Compute summary stats
    const totalFiles   = _allItems.filter(i => i.type === "file").length;
    const totalFolders = _allItems.filter(i => i.type === "folder").length;
    const totalSize    = _allItems.reduce((sum, i) => sum + (i.type === "file" ? i.size : 0), 0);

    let html = `<div class="report-summary">
      <div class="report-stat"><span class="report-stat-label">Files</span><span class="report-stat-value">${totalFiles.toLocaleString()}</span></div>
      <div class="report-stat"><span class="report-stat-label">Folders</span><span class="report-stat-value">${totalFolders.toLocaleString()}</span></div>
      <div class="report-stat"><span class="report-stat-label">Total size</span><span class="report-stat-value">${formatFileSize(totalSize)}</span></div>
      <div class="report-stat"><span class="report-stat-label">Location</span><span class="report-stat-value">${_esc(_currentPrefix || "/")}</span></div>
    </div>`;

    // Filter bar
    html += `<div class="report-filter-bar">
      <select class="report-filter-select" id="reportTypeFilter" title="Filter by type">
        <option value="">All types</option>
        <option value="file"${_filterType === "file" ? " selected" : ""}>Files only</option>
        <option value="folder"${_filterType === "folder" ? " selected" : ""}>Folders only</option>
      </select>
      <input type="text" class="report-filter-input" id="reportTextFilter" placeholder="Filter name, path, content type\u2026" value="${_esc(_filterText)}" />
    </div>`;

    if (items.length === 0) {
      html += `<div class="report-empty">No items match the current filter.</div>`;
    } else {
      html += `<table class="report-table"><thead><tr>
        <th>Type</th>
        <th>Name</th>
        <th>Full Path</th>
        <th>Path Length</th>
        <th>Blob URL</th>
        <th>Size</th>
        <th>Size (bytes)</th>
        <th>Content Type</th>
        <th>Last Modified</th>
        <th>Created On</th>
        <th>ETag</th>
        <th>MD5</th>
      </tr></thead><tbody>`;

      const { accountName, containerName } = CONFIG.storage;
      for (const item of items) {
        const typeClass = item.type === "folder" ? "report-type-folder" : "report-type-file";
        const icon      = item.type === "folder" ? "\uD83D\uDCC1" : "\uD83D\uDCC4";
        const size      = item.type === "file" ? formatFileSize(item.size) : "\u2014";
        const sizeBytes = item.type === "file" ? item.size.toLocaleString() : "\u2014";
        const ct        = _esc(item.contentType || "\u2014");
        const modified  = item.lastModified ? new Date(item.lastModified).toLocaleString() : "\u2014";
        const created   = item.createdOn ? new Date(item.createdOn).toLocaleString() : "\u2014";
        const encoded   = item.name.split("/").map(encodeURIComponent).join("/");
        const blobUrl   = `https://${accountName}.blob.core.windows.net/${containerName}/${encoded}`;
        const pathLen   = item.name.length;

        html += `<tr>
          <td class="report-col-type"><span class="report-col-type-label ${typeClass}">${icon} ${item.type}</span></td>
          <td class="report-col-name" title="${_esc(item.displayName)}">${_esc(item.displayName)}</td>
          <td class="report-col-path" title="${_esc(item.name)}">${_esc(item.name)}</td>
          <td class="report-col-num">${pathLen}</td>
          <td class="report-col-url" title="${_esc(blobUrl)}"><a href="${_esc(blobUrl)}" target="_blank" rel="noopener">${_esc(blobUrl)}</a></td>
          <td class="report-col-size">${size}</td>
          <td class="report-col-num">${sizeBytes}</td>
          <td title="${ct}">${ct}</td>
          <td style="white-space:nowrap">${_esc(modified)}</td>
          <td style="white-space:nowrap">${_esc(created)}</td>
          <td class="report-col-mono">${_esc(item.etag || "\u2014")}</td>
          <td class="report-col-mono">${_esc(item.md5 || "\u2014")}</td>
        </tr>`;
      }
      html += `</tbody></table>`;
    }

    body.innerHTML = html;
    metaEl.textContent = `${items.length} of ${_allItems.length} items`;

    // Wire filter controls
    const typeSel = _el("reportTypeFilter");
    const textInp = _el("reportTextFilter");
    if (typeSel) typeSel.onchange = () => { _filterType = typeSel.value; _renderReportTable(); };
    if (textInp) {
      let _debounce;
      textInp.oninput = () => {
        clearTimeout(_debounce);
        _debounce = setTimeout(() => { _filterText = textInp.value; _renderReportTable(); }, 250);
      };
    }
  }
}

// ── Download ────────────────────────────────────────────

async function _handleDownload(file) {
  _showLoading(true);
  try {
    await downloadBlob(file.name);
    _audit("download", file.name);
  } catch (err) {
    console.error("[app] Download error:", err);
    _showError(`Download failed: ${err.message}`);
  } finally {
    _showLoading(false);
  }
}

async function _handleFolderDownload(folder) {
  _showLoading(true);
  _showToast(`⏳ Preparing "${folder.displayName}.zip"…`);
  try {
    await downloadFolderAsZip(folder.name, folder.displayName);
    _audit("download", folder.name, { type: "folder" });
    _showToast(`✅ "${folder.displayName}.zip" downloaded`);
  } catch (err) {
    console.error("[app] Folder download error:", err);
    _showError(`Folder download failed: ${err.message}`);
  } finally {
    _showLoading(false);
  }
}

async function _downloadCurrentLevel() {
  const { containerName } = CONFIG.storage;
  const displayName = _currentPrefix
    ? _currentPrefix.replace(/\/$/, "").split("/").pop()
    : containerName;
  _showToast(`⏳ Scanning files…`, 300000);
  try {
    await downloadFolderAsZip(_currentPrefix, displayName, (done, total) => {
      const pct = Math.round((done / total) * 100);
      _updateToastMsg(`⏳ Preparing ${done}/${total} files — ${pct}%`);
    });
    _audit("download", _currentPrefix || "/", { type: "folder" });
    _showToast(`✅ "${displayName}.zip" downloaded`);
  } catch (err) {
    console.error("[app] Download error:", err);
    _showError(`Download failed: ${err.message}`);
  }
}

// ── Upload panel ─────────────────────────────────────────────
// ── Properties modal ────────────────────────────────────────

async function _showFolderProperties(folder) {
  const modal = _el("propsModal");
  const table = _el("propsTable");
  const title = _el("propsModalTitle");
  title.textContent = `Properties — ${folder.displayName}`;
  table.innerHTML = "<tr><td colspan='2'>Scanning folder…</td></tr>";
  modal.classList.remove("hidden");

  try {
    const stats = await getFolderStats(folder.name);
    const rows = [
      ["Name",            folder.displayName],
      ["Full path",       folder.name],
      ["Subfolders",      stats.totalFolders.toLocaleString()],
      ["Files",           stats.totalFiles.toLocaleString()],
      ["Total size",      formatFileSize(stats.totalSize)],
    ];
    table.innerHTML = rows
      .map(([k, v]) => `<tr><td>${_esc(k)}</td><td>${_esc(v)}</td></tr>`)
      .join("");
  } catch (err) {
    table.innerHTML = `<tr><td colspan='2' style='color:var(--error)'>Failed to scan folder: ${_esc(err.message)}</td></tr>`;
  }

  _el("propsModalClose").onclick = () => modal.classList.add("hidden");
  modal.onclick = (e) => { if (e.target === modal) modal.classList.add("hidden"); };
}

async function _showProperties(file) {
  const modal = _el("propsModal");
  const table = _el("propsTable");
  const title = _el("propsModalTitle");
  title.textContent = `Properties — ${file.displayName}`;
  table.innerHTML = "<tr><td colspan='2'>Loading…</td></tr>";
  modal.classList.remove("hidden");

  try {
    const props = await getBlobProperties(file.name);
    const { accountName, containerName } = CONFIG.storage;
    const blobUrl = `https://${accountName}.blob.core.windows.net/${containerName}/${_encodePath(file.name)}`;

    const rows = [
      ["Name",          file.name],
      ["Display name",  file.displayName],
      ["URL",           blobUrl],
      ["Size",          formatFileSize(file.size)],
      ["Content type",  props["content-type"]  || file.contentType || "—"],
      ["Last modified", props["last-modified"]  || formatDate(file.lastModified)],
      ["Created",       props["x-ms-creation-time"] || formatDate(file.createdOn) || "—"],
      ["ETag",          props["etag"]           || file.etag || "—"],
      ["Content-MD5",   props["content-md5"]    || file.md5  || "—"],
      ["Blob type",     props["x-ms-blob-type"] || "—"],
      ["Lease status",  props["x-ms-lease-status"] || "—"],
    ];

    // Also include any x-ms-meta-* user metadata
    Object.entries(props)
      .filter(([k]) => k.startsWith("x-ms-meta-"))
      .forEach(([k, v]) => rows.push([`Metadata: ${k.replace("x-ms-meta-", "")}`, v]));

    table.innerHTML = rows
      .map(([k, v]) => `<tr><td>${_esc(k)}</td><td>${_esc(v)}</td></tr>`)
      .join("");
  } catch (err) {
    table.innerHTML = `<tr><td colspan='2' style='color:var(--error)'>Failed to load properties: ${_esc(err.message)}</td></tr>`;
  }

  _el("propsModalClose").onclick = () => modal.classList.add("hidden");
  modal.onclick = (e) => { if (e.target === modal) modal.classList.add("hidden"); };
}

// ── Copy URL ─────────────────────────────────────────────────────────

// ── Copy URL menu ───────────────────────────────────────────────

let _copyMenuOpen = null;

function _showCopyMenu(btn, item, kind, pos) {
  // Close any already-open menu
  if (_copyMenuOpen) { _copyMenuOpen.remove(); _copyMenuOpen = null; }

  const { accountName, containerName } = CONFIG.storage;
  const encoded = item.name.split("/").map(encodeURIComponent).join("/");
  const blobUrl = `https://${accountName}.blob.core.windows.net/${containerName}/${encoded}`;
  const appUrl  = _buildAppUrl(item.name);

  const menu = document.createElement("div");
  menu.className = "copy-menu";

  // ── Helper: build a menu item button ──
  function _menuBtn(icon, label, sub, onClick) {
    const mi = document.createElement("button");
    mi.type = "button";
    mi.className = "copy-menu-item";
    const iconSpan = document.createElement("span");
    iconSpan.className = "copy-menu-icon";
    iconSpan.textContent = icon;
    const textSpan = document.createElement("span");
    const strong = document.createElement("strong");
    strong.textContent = label;
    const small = document.createElement("small");
    small.textContent = sub;
    textSpan.appendChild(strong);
    textSpan.appendChild(small);
    mi.appendChild(iconSpan);
    mi.appendChild(textSpan);
    mi.addEventListener("click", (e) => {
      e.stopPropagation();
      menu.remove();
      _copyMenuOpen = null;
      onClick();
    });
    return mi;
  }

  // ── Copy items ──
  menu.appendChild(_menuBtn("🗄️", "Blob URL",
    "Direct link to the file in Azure Storage",
    () => _copyToClipboard(blobUrl)));

  menu.appendChild(_menuBtn("🔗", "App link",
    "Link that opens this explorer at this location",
    () => {
      if (CONFIG.app.allowDownload) {
        const dlFn = kind === "folder"
          ? () => _handleFolderDownload(item)
          : () => _handleDownload(item);
        const doToast = () => _showToast("📋 App link copied!", 8000, "⬇ Download", dlFn);
        if (navigator.clipboard && window.isSecureContext) {
          navigator.clipboard.writeText(appUrl).then(doToast).catch(() => { _fallbackCopy(appUrl); doToast(); });
        } else {
          _fallbackCopy(appUrl);
          doToast();
        }
      } else {
        _copyToClipboard(appUrl);
      }
    }));

  // ── Email items — only when signed in with OAuth ──
  if (_canEmail()) {
    const divider = document.createElement("div");
    divider.className = "copy-menu-divider";
    menu.appendChild(divider);

    const itemName = item.displayName || item.name.split("/").filter(Boolean).pop() || item.name;

    menu.appendChild(_menuBtn("📧", "Email blob URL",
      "Send the direct Azure Storage link via email",
      () => _showEmailComposeModal(itemName, blobUrl)));

    menu.appendChild(_menuBtn("📧", "Email app link",
      "Send the explorer link via email",
      () => _showEmailComposeModal(itemName, appUrl)));
  }

  document.body.appendChild(menu);
  _copyMenuOpen = menu;

  // Position: use fixed coords if provided, otherwise anchor to the button
  if (pos) {
    const menuW = menu.offsetWidth || 260;
    const menuH = menu.offsetHeight || 200;
    let x = pos.x;
    let y = pos.y;
    if (x + menuW > window.innerWidth - 8)  x = window.innerWidth - menuW - 8;
    if (y + menuH > window.innerHeight - 8) y = window.innerHeight - menuH - 8;
    if (x < 8) x = 8;
    if (y < 8) y = 8;
    menu.style.top  = `${y}px`;
    menu.style.left = `${x}px`;
  } else {
    // Use fixed positioning — no scroll offset math needed
    const rect = btn.getBoundingClientRect();
    const menuW = menu.offsetWidth || 260;
    let left = rect.left;
    if (left + menuW > window.innerWidth - 8) left = window.innerWidth - menuW - 8;
    menu.style.top  = `${rect.bottom + 4}px`;
    menu.style.left = `${Math.max(8, left)}px`;
  }

  // Dismiss on outside click
  const dismiss = (e) => {
    if (!menu.contains(e.target)) {
      menu.remove();
      _copyMenuOpen = null;
      document.removeEventListener("click", dismiss, true);
    }
  };
  setTimeout(() => document.addEventListener("click", dismiss, true), 0);
}

function _copyToClipboard(text) {
  if (navigator.clipboard && window.isSecureContext) {
    navigator.clipboard.writeText(text)
      .then(() => _showToast("📋 Copied to clipboard"))
      .catch(() => {
        if (_fallbackCopy(text)) _showToast("📋 Copied to clipboard");
        else _showToast("Copy failed — please copy manually");
      });
  } else {
    if (_fallbackCopy(text)) _showToast("📋 Copied to clipboard");
    else _showToast("Copy failed — please copy manually");
  }
}

function _fallbackCopy(text) {
  const ta = document.createElement("textarea");
  ta.value = text;
  ta.style.cssText = "position:fixed;top:-9999px;left:-9999px;opacity:0";
  document.body.appendChild(ta);
  ta.focus();
  ta.select();
  let ok = false;
  try { document.execCommand("copy"); ok = true; } catch {}
  document.body.removeChild(ta);
  return ok;
}

// ── Email compose modal ──────────────────────────────────────────

/**
 * Open the email compose modal pre-filled with a link.
 * @param {string} itemName  Display name of the blob/folder
 * @param {string} url       The URL to include in the email body
 */
function _showEmailComposeModal(itemName, url) {
  const modal     = _el("emailModal");
  const toInput   = _el("emailTo");
  const subInput  = _el("emailSubject");
  const bodyInput = _el("emailBody");
  const errEl     = _el("emailError");
  const sendingEl = _el("emailSending");
  const closeBtn  = _el("emailModalClose");
  const cancelBtn = _el("emailCancelBtn");
  const sendBtn   = _el("emailSendBtn");

  const user = getUser();
  const senderName = user ? user.name : "Someone";

  // Pre-fill
  toInput.value   = "";
  subInput.value  = `${senderName} shared "${itemName}" with you`;
  bodyInput.value = `Hi,\n\n${senderName} shared a file with you from Azure Blob Storage:\n\n${url}\n\nRegards,\n${senderName}`;
  errEl.textContent = "";
  errEl.classList.add("hidden");
  sendingEl.classList.add("hidden");
  sendBtn.disabled = false;
  modal.classList.remove("hidden");
  toInput.focus();

  const close = () => { modal.classList.add("hidden"); };

  closeBtn.onclick  = close;
  cancelBtn.onclick = close;
  modal.addEventListener("click", (e) => { if (e.target === modal) close(); });

  sendBtn.onclick = async () => {
    errEl.classList.add("hidden");
    const rawTo = toInput.value.trim();
    if (!rawTo) {
      errEl.textContent = "Please enter at least one recipient email address.";
      errEl.classList.remove("hidden");
      toInput.focus();
      return;
    }

    // Parse recipients — split on ; , or space
    const recipients = rawTo.split(/[;,\s]+/).map(e => e.trim()).filter(Boolean);
    const emailRe = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    const invalid = recipients.filter(r => !emailRe.test(r));
    if (invalid.length) {
      errEl.textContent = `Invalid email address: ${invalid.join(", ")}`;
      errEl.classList.remove("hidden");
      return;
    }

    const subject = subInput.value.trim() || "(no subject)";
    const body    = bodyInput.value;

    sendBtn.disabled = true;
    sendingEl.classList.remove("hidden");

    try {
      await _sendEmailViaGraph(recipients, subject, body);
      close();
      _showToast("📧 Email sent successfully!");
    } catch (err) {
      errEl.textContent = err.message || "Failed to send email.";
      errEl.classList.remove("hidden");
    } finally {
      sendBtn.disabled = false;
      sendingEl.classList.add("hidden");
    }
  };
}

/**
 * Send an email using the Microsoft Graph /me/sendMail API.
 * Requires an Exchange Online license and the Mail.Send scope.
 * @param {string[]} recipients  Array of email addresses
 * @param {string}   subject
 * @param {string}   body        Plain text body
 */
async function _sendEmailViaGraph(recipients, subject, body) {
  const token = await getGraphToken();

  const payload = {
    message: {
      subject,
      body: {
        contentType: "Text",
        content: body,
      },
      toRecipients: recipients.map(addr => ({
        emailAddress: { address: addr },
      })),
    },
    saveToSentItems: true,
  };

  const res = await fetch("https://graph.microsoft.com/v1.0/me/sendMail", {
    method:  "POST",
    headers: {
      Authorization:  `Bearer ${token}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify(payload),
  });

  if (!res.ok) {
    let msg = `Failed to send email (${res.status})`;
    try {
      const data = await res.json();
      if (data?.error?.message) {
        msg = data.error.message;
        // Friendly message for missing license
        if (/MailboxNotEnabledForRESTAPI|MailboxNotFound/i.test(data.error.code || "")) {
          msg = "Your account does not have an Exchange Online mailbox. Please ask your admin to assign a license.";
        }
      }
    } catch {}
    throw new Error(msg);
  }
}

// ── Rename modal ─────────────────────────────────────────────────

function _showRenameModal(item, type) {
  const modal    = _el("renameModal");
  const input    = _el("renameInput");
  const errEl    = _el("renameError");
  const closeBtn = _el("renameModalClose");
  const cancelBtn = _el("renameCancelBtn");
  const confirmBtn = _el("renameConfirmBtn");

  _el("renameModalTitle").textContent = `Rename ${type === "folder" ? "folder" : "file"}`;
  input.value = item.displayName;
  errEl.textContent = "";
  errEl.classList.add("hidden");
  modal.classList.remove("hidden");
  input.select();

  const close = () => modal.classList.add("hidden");

  const doRename = async () => {
    const newName = input.value.trim();
    if (!newName || newName === item.displayName) { close(); return; }
    if (newName.includes("/")) {
      errEl.textContent = "Name cannot contain \"/\"";
      errEl.classList.remove("hidden");
      return;
    }

    // Build full blob paths (folder items end with /)
    const parentPrefix = _currentPrefix;
    const srcName  = item.name;
    const destName = type === "folder"
      ? parentPrefix + newName + "/"
      : parentPrefix + newName;

    confirmBtn.disabled = true;
    confirmBtn.textContent = "Renaming…";
    errEl.classList.add("hidden");
    try {
      if (type === "folder") {
        // Recursively collect ALL blobs under this folder (including nested subfolders)
        async function _collectAllBlobs(prefix) {
          const { folders, files } = await listBlobsAtPrefix(prefix);
          let blobs = files.map(f => f.name);
          for (const sub of folders) blobs = blobs.concat(await _collectAllBlobs(sub.name));
          return blobs;
        }
        const allBlobs = await _collectAllBlobs(srcName);
        // If folder is empty just reflect the new name (virtual folders don't exist as blobs)
        if (allBlobs.length === 0) {
          // No real blobs to move — just close and refresh
        } else {
          // Guard against excessively large operations that could hang the browser.
          const MAX_RENAME_ITEMS = 5000;
          if (allBlobs.length > MAX_RENAME_ITEMS) {
            throw new Error(
              `Folder is too large to rename in a single operation (${allBlobs.length} items). ` +
              `Please move or rename smaller subsets.`
            );
          }

          // Process renames in small batches with short delays to keep the UI responsive.
          const BATCH_SIZE = 100;
          const BATCH_DELAY_MS = 10;

          for (let i = 0; i < allBlobs.length; i += BATCH_SIZE) {
            const batch = allBlobs.slice(i, i + BATCH_SIZE);
            await Promise.all(
              batch.map(blob => {
                const rel  = blob.slice(srcName.length);
                const dest = destName + rel;
                return renameBlob(blob, dest);
              })
            );

            // Yield to the event loop between batches to avoid long blocking periods.
            if (i + BATCH_SIZE < allBlobs.length) {
              await new Promise(resolve => setTimeout(resolve, BATCH_DELAY_MS));
            }
          }
        }
      } else {
        await renameBlob(srcName, destName);
      }
      _audit("rename", srcName, { newName: destName });
      close();
      _loadFiles(_currentPrefix);
      _showToast(`✓ Renamed to "${newName}"`);
    } catch (err) {
      errEl.textContent = err.message;
      errEl.classList.remove("hidden");
    } finally {
      confirmBtn.disabled = false;
      confirmBtn.textContent = "Rename";
    }
  };

  // Wire buttons — replace listeners cleanly
  confirmBtn.onclick  = doRename;
  cancelBtn.onclick   = close;
  closeBtn.onclick    = close;
  modal.onclick = (e) => { if (e.target === modal) close(); };
  input.onkeydown = (e) => { if (e.key === "Enter") doRename(); if (e.key === "Escape") close(); };
}

// ── Copy / Move modal ────────────────────────────────────────────

function _showCopyModal(item, type) { _showCopyMoveModal(item, type, "copy"); }
function _showMoveModal(item, type) { _showCopyMoveModal(item, type, "move"); }

/**
 * Generic modal that lets the user pick a destination folder and then
 * copies or moves one or more items there.
 * @param {object|null}    item   Single item (or null for bulk mode)
 * @param {string}         type   "file" | "folder" | "bulk"
 * @param {"copy"|"move"}  mode
 */
function _showCopyMoveModal(item, type, mode) {
  const isMove = mode === "move";
  const verb   = isMove ? "Move" : "Copy";
  const icon   = isMove ? "📦" : "📋";

  const modal      = _el("copyModal");
  const treeEl     = _el("copyTreeContainer");
  const destInput  = _el("copyDestInput");
  const errEl      = _el("copyError");
  const closeBtn   = _el("copyModalClose");
  const cancelBtn  = _el("copyCancelBtn");
  const confirmBtn = _el("copyConfirmBtn");

  // Title
  if (type === "bulk") {
    _el("copyModalTitle").textContent = `${verb} ${_selection.size} item${_selection.size !== 1 ? "s" : ""} to…`;
  } else {
    const label = type === "folder" ? "folder" : "file";
    _el("copyModalTitle").textContent = `${verb} ${label} — ${item.displayName}`;
  }
  destInput.value = "";
  errEl.textContent = "";
  errEl.classList.add("hidden");
  treeEl.innerHTML = "";
  confirmBtn.textContent = `${icon} ${verb}`;
  modal.classList.remove("hidden");

  // ── Build a mini folder-picker tree ──────────────────────────
  const containerName = CONFIG.storage.containerName;
  let _selectedPrefix = "";

  function selectPrefix(prefix, row) {
    _selectedPrefix = prefix;
    destInput.value = prefix;
    treeEl.querySelectorAll(".copy-tree-active").forEach(r => r.classList.remove("copy-tree-active"));
    if (row) row.classList.add("copy-tree-active");
  }

  function makeNode(displayName, prefix) {
    const node = document.createElement("div");
    node.className = "copy-tree-node";
    node.dataset.prefix = prefix;

    const row = document.createElement("div");
    row.className = "copy-tree-row";

    const chevron = document.createElement("button");
    chevron.type = "button";
    chevron.className = "copy-tree-chevron";
    chevron.innerHTML = "&#9654;";

    const icon = document.createElement("span");
    icon.textContent = prefix === "" ? "📦" : "📁";

    const lbl = document.createElement("span");
    lbl.textContent = displayName;

    row.appendChild(chevron);
    row.appendChild(icon);
    row.appendChild(lbl);

    const childrenEl = document.createElement("div");
    childrenEl.className = "copy-tree-children";

    node.appendChild(row);
    node.appendChild(childrenEl);

    let loaded = false;

    async function ensureLoaded() {
      if (loaded) return;
      loaded = true;
      childrenEl.innerHTML = '<div style="font-size:12px;color:var(--text-muted);padding:2px 6px">Loading…</div>';
      try {
        const { folders } = await listBlobsAtPrefix(prefix);
        childrenEl.innerHTML = "";
        if (folders.length === 0) {
          chevron.style.visibility = "hidden";
        } else {
          for (const f of folders) childrenEl.appendChild(makeNode(f.displayName, f.name));
        }
      } catch {
        childrenEl.innerHTML = '<div style="font-size:12px;color:var(--error);padding:2px 6px">Failed</div>';
      }
    }

    chevron.addEventListener("click", async (e) => {
      e.stopPropagation();
      const expanding = !node.classList.contains("ct-expanded");
      node.classList.toggle("ct-expanded");
      if (expanding) await ensureLoaded();
    });

    row.addEventListener("click", async () => {
      selectPrefix(prefix, row);
      if (!node.classList.contains("ct-expanded")) {
        node.classList.add("ct-expanded");
        await ensureLoaded();
      }
    });

    return node;
  }

  // Root node
  const root = makeNode(containerName, "");
  root.classList.add("ct-expanded");
  treeEl.appendChild(root);
  // Eagerly load root children
  (async () => {
    const rootChildren = root.querySelector(".copy-tree-children");
    rootChildren.innerHTML = '<div style="font-size:12px;color:var(--text-muted);padding:2px 6px">Loading…</div>';
    try {
      const { folders } = await listBlobsAtPrefix("");
      rootChildren.innerHTML = "";
      for (const f of folders) rootChildren.appendChild(makeNode(f.displayName, f.name));
    } catch {
      rootChildren.innerHTML = '<div style="font-size:12px;color:var(--error);padding:2px 6px">Failed</div>';
    }
    // Mark root loaded
    root._loaded = true;
  })();

  // Select root by default
  selectPrefix("", root.querySelector(".copy-tree-row"));

  const close = () => modal.classList.add("hidden");

  // Helper — prompt the user about a single conflict, returns "skip" | "overwrite" | "overwriteAll"
  function _promptCopyConflict(displayName, destPath, itemType) {
    return new Promise((resolve) => {
      const cfModal   = _el("copyConflictModal");
      const msgEl     = _el("copyConflictMsg");
      const closeBtn  = _el("copyConflictClose");
      const skipBtn   = _el("copyConflictSkipBtn");
      const owBtn     = _el("copyConflictOverwriteBtn");
      const owAllBtn  = _el("copyConflictOverwriteAllBtn");

      const kind = itemType === "folder" ? "folder" : "file";
      msgEl.innerHTML =
        `The ${kind} <strong>${_esc(displayName)}</strong> already exists at ` +
        `<strong>${_esc(destPath || "/")}</strong>.<br>Do you want to overwrite it?`;
      cfModal.classList.remove("hidden");

      const done = (result) => { cfModal.classList.add("hidden"); resolve(result); };
      skipBtn.onclick   = () => done("skip");
      owBtn.onclick     = () => done("overwrite");
      owAllBtn.onclick  = () => done("overwriteAll");
      closeBtn.onclick  = () => done("skip");
      cfModal.onclick   = (e) => { if (e.target === cfModal) done("skip"); };
    });
  }

  // Track overwrite-all state across the current operation
  let _overwriteAll = false;

  // Helper — check if destination exists, prompt if needed. Returns true to proceed, false to skip.
  async function _shouldProceed(displayName, destBlobOrPrefix, normDest, itemType) {
    if (_overwriteAll) return true;

    let exists = false;
    if (itemType === "folder") {
      // Check if destination folder has any content
      const { folders, files } = await listBlobsAtPrefix(normDest + displayName + "/");
      exists = folders.length > 0 || files.length > 0;
    } else {
      exists = await blobExists(destBlobOrPrefix);
    }

    if (!exists) return true;

    const result = await _promptCopyConflict(displayName, normDest, itemType);
    if (result === "overwriteAll") { _overwriteAll = true; return true; }
    if (result === "overwrite")    return true;
    return false; // "skip"
  }

  // Helper — recursively collect all blobs under a folder prefix
  async function _collectAll(prefix) {
    const { folders, files } = await listBlobsAtPrefix(prefix);
    let blobs = files.map(f => f.name);
    for (const sub of folders) blobs = blobs.concat(await _collectAll(sub.name));
    return blobs;
  }

  // Helper — copy (and optionally delete) a single item
  async function _processItem(srcItem, srcType, normDest) {
    if (srcType === "folder") {
      const allBlobs = await _collectAll(srcItem.name);
      if (allBlobs.length === 0) {
        const placeholder = new File([""], ".keep", { type: "application/octet-stream" });
        await uploadBlob(normDest + srcItem.displayName + "/.keep", placeholder, null, {});
      } else {
        const MAX = 5000;
        if (allBlobs.length > MAX) throw new Error(`Folder "${srcItem.displayName}" too large (${allBlobs.length} items).`);
        const BATCH = 100;
        for (let i = 0; i < allBlobs.length; i += BATCH) {
          await Promise.all(allBlobs.slice(i, i + BATCH).map(blob => {
            const rel  = blob.slice(srcItem.name.length);
            return copyBlobOnly(blob, normDest + srcItem.displayName + "/" + rel);
          }));
          if (i + BATCH < allBlobs.length) await new Promise(r => setTimeout(r, 10));
        }
      }
      if (isMove) await deleteFolderContents(srcItem.name);
    } else {
      await copyBlobOnly(srcItem.name, normDest + srcItem.displayName);
      if (isMove) await deleteBlob(srcItem.name);
    }
  }

  // Helper — determine the parent prefix of an item
  function _parentOf(name, itemType) {
    if (itemType === "folder") {
      // e.g. "foo/bar/" → parent is "foo/"
      const trimmed = name.slice(0, -1); // remove trailing /
      const idx = trimmed.lastIndexOf("/");
      return idx >= 0 ? trimmed.slice(0, idx + 1) : "";
    }
    // file: "foo/bar/file.txt" → "foo/bar/"
    const idx = name.lastIndexOf("/");
    return idx >= 0 ? name.slice(0, idx + 1) : "";
  }

  const doAction = async () => {
    const destPrefix = destInput.value.trim();
    const normDest   = destPrefix && !destPrefix.endsWith("/") ? destPrefix + "/" : destPrefix;

    // Validation — prevent moving/copying a folder into itself
    if (type === "folder" && item && normDest.startsWith(item.name)) {
      errEl.textContent = `Cannot ${verb.toLowerCase()} a folder into itself or a subfolder of itself.`;
      errEl.classList.remove("hidden");
      return;
    }

    // Bulk validation — same check for every selected folder
    if (type === "bulk") {
      for (const name of _selection) {
        if (name.endsWith("/") && normDest.startsWith(name)) {
          errEl.textContent = `Cannot ${verb.toLowerCase()} a folder into itself or a subfolder of itself.`;
          errEl.classList.remove("hidden");
          return;
        }
      }
    }

    confirmBtn.disabled = true;
    confirmBtn.textContent = `${verb === "Move" ? "Moving" : "Copying"}…`;
    errEl.classList.add("hidden");

    // Reset overwrite-all state for this operation
    _overwriteAll = false;

    try {
      let bulkCount = 0;
      const skipped = [];
      let conflictSkipped = 0;
      let selSnap = [];

      if (type === "bulk") {
        bulkCount = _selection.size;
        selSnap = [..._selection];             // snapshot before clear
        // Process each selected item, skipping same-location items
        for (const name of selSnap) {
          const isFolder = name.endsWith("/");
          const displayName = isFolder
            ? name.slice(0, -1).split("/").pop()
            : name.split("/").pop();
          const parentPfx = _parentOf(name, isFolder ? "folder" : "file");
          if (parentPfx === normDest) {
            skipped.push(displayName);
            continue;
          }
          const itemType = isFolder ? "folder" : "file";
          const destBlob = isFolder ? "" : normDest + displayName;
          const proceed = await _shouldProceed(displayName, destBlob, normDest, itemType);
          if (!proceed) { conflictSkipped++; continue; }
          await _processItem({ name, displayName }, itemType, normDest);
        }
        _selection.clear();
        _updateSelectionBar();
      } else {
        // Single item — check same-location
        const parentPfx = _parentOf(item.name, type);
        if (parentPfx === normDest) {
          close();
          _showToast(`⏭ Skipped "${item.displayName}" — already in ${normDest || "/"}`);
          return;
        }
        const destBlob = type === "folder" ? "" : normDest + item.displayName;
        const proceed = await _shouldProceed(item.displayName, destBlob, normDest, type);
        if (!proceed) {
          close();
          _showToast(`⏭ Skipped "${item.displayName}" — not overwritten`);
          return;
        }
        await _processItem(item, type, normDest);
      }

      // Audit: log each processed item
      if (type === "bulk") {
        for (const name of selSnap) {
          _audit(isMove ? "move" : "copy", name, { destination: normDest });
        }
      } else {
        _audit(isMove ? "move" : "copy", item.name, { destination: normDest });
      }

      close();
      _invalidateTreeChildren(normDest);
      if (isMove) _invalidateTreeChildren(_currentPrefix);
      _loadFiles(_currentPrefix);

      const totalSkipped = skipped.length + conflictSkipped;
      const processed = type === "bulk" ? bulkCount - totalSkipped : 1;
      const itemLabel = type === "bulk"
        ? `${processed} item${processed !== 1 ? "s" : ""}`
        : `"${item.displayName}"`;

      if (totalSkipped > 0 && processed > 0) {
        _showToast(`✓ ${verb === "Move" ? "Moved" : "Copied"} ${itemLabel} → ${normDest || "/"} (skipped ${totalSkipped})`);
      } else if (totalSkipped > 0 && processed === 0) {
        _showToast(`⏭ All ${totalSkipped} item${totalSkipped !== 1 ? "s" : ""} skipped`);
      } else {
        _showToast(`✓ ${verb === "Move" ? "Moved" : "Copied"} ${itemLabel} → ${normDest || "/"}`);
      }
    } catch (err) {
      errEl.textContent = err.message;
      errEl.classList.remove("hidden");
    } finally {
      confirmBtn.disabled = false;
      confirmBtn.textContent = `${icon} ${verb}`;
    }
  };

  confirmBtn.onclick = doAction;
  cancelBtn.onclick  = close;
  closeBtn.onclick   = close;
  modal.onclick = (e) => { if (e.target === modal) close(); };
  destInput.onkeydown = (e) => { if (e.key === "Enter") doAction(); if (e.key === "Escape") close(); };
}

// ── Toast helper ──────────────────────────────────────────────────

// ── Selection bar ───────────────────────────────────────────────

function _updateSelectionBar() {
  const bar   = _el("selectionBar");
  const count = _selection.size;
  bar.classList.toggle("hidden", count === 0);
  _el("selectionCount").textContent =
    `${count} item${count !== 1 ? "s" : ""} selected`;
  // Only show delete button when user can upload (contributor+)
  _el("selBulkDeleteBtn").classList.toggle("hidden", !_canDeleteItems());
  _el("selBulkCopyBtn").classList.toggle("hidden", !_canCopyItems());
  _el("selBulkMoveBtn").classList.toggle("hidden", !_canMoveItems());
}

// Wire selection bar buttons once (they are always in the DOM)
document.addEventListener("DOMContentLoaded", () => {
  _el("selClearBtn").addEventListener("click", () => {
    _selection.clear();
    _updateSelectionBar();
    // Uncheck all visible checkboxes
    document.querySelectorAll(".row-chk, .card-chk").forEach(c => c.checked = false);
    document.querySelectorAll(".row-selected, .card-selected").forEach(el => {
      el.classList.remove("row-selected", "card-selected");
    });
    const all = _el("selectAllChk");
    if (all) all.checked = false;
  });
  _el("selBulkDownloadBtn").addEventListener("click", _bulkDownload);
  _el("selBulkCopyBtn").addEventListener("click",     () => _showCopyMoveModal(null, "bulk", "copy"));
  _el("selBulkMoveBtn").addEventListener("click",     () => _showCopyMoveModal(null, "bulk", "move"));
  _el("selBulkDeleteBtn").addEventListener("click",   _bulkDelete);
});

async function _bulkDownload() {
  if (_selection.size === 0) return;
  _showLoading(true);
  _showToast(`⏳ Building ZIP for ${_selection.size} item${_selection.size !== 1 ? "s" : ""}…`);
  try {
    const { accountName, containerName } = CONFIG.storage;

    // Recursively collect all blob names for a name (file or folder)
    async function collectNames(name) {
      if (name.endsWith("/")) {
        const { folders, files } = await listBlobsAtPrefix(name);
        let names = files.map(f => f.name);
        const sub = await Promise.all(folders.map(f => collectNames(f.name)));
        for (const s of sub) names = names.concat(s);
        return names;
      }
      return [name];
    }

    let allNames = [];
    for (const name of _selection) {
      allNames = allNames.concat(await collectNames(name));
    }

    if (allNames.length === 0) throw new Error("Selection contains no downloadable files.");

    // Determine zip entry paths — strip common prefix if all items share one
    const stripPrefix = _currentPrefix;

    // Fetch blobs in parallel batches for speed
    const FETCH_BATCH = 6;
    const entries = [];
    for (let i = 0; i < allNames.length; i += FETCH_BATCH) {
      const batch = allNames.slice(i, i + FETCH_BATCH);
      const results = await Promise.all(batch.map(async (name) => {
        const authHeaders = await _storageAuthHeaders();
        const rawUrl = `https://${accountName}.blob.core.windows.net/${containerName}/${_encodePath(name)}`;
        const url = _sasUrl(rawUrl);
        const res = await fetch(url, { headers: authHeaders });
        if (!res.ok) throw new Error(`Failed to fetch "${name}" (${res.status})`);
        const data = new Uint8Array(await res.arrayBuffer());
        return { name: name.startsWith(stripPrefix) ? name.slice(stripPrefix.length) : name, data };
      }));
      entries.push(...results);
    }

    const zipBlob = _buildZip(entries);
    const objUrl  = URL.createObjectURL(zipBlob);
    const anchor  = document.createElement("a");
    anchor.href     = objUrl;
    anchor.download = `selection.zip`;
    document.body.appendChild(anchor);
    anchor.click();
    document.body.removeChild(anchor);
    URL.revokeObjectURL(objUrl);
    _showToast(`✅ selection.zip downloaded (${entries.length} file${entries.length !== 1 ? "s" : ""})`);
  } catch (err) {
    console.error("[app] Bulk download error:", err);
    _showError(`Bulk download failed: ${err.message}`);
  } finally {
    _showLoading(false);
  }
}

async function _bulkDelete() {
  if (_selection.size === 0 || !_canDeleteItems()) return;
  const count = _selection.size;
  const modal      = _el("deleteModal");
  const msgEl      = _el("deleteModalMsg");
  const closeBtn   = _el("deleteModalClose");
  const cancelBtn  = _el("deleteCancelBtn");
  const confirmBtn = _el("deleteConfirmBtn");

  msgEl.textContent = `Are you sure you want to permanently delete ${count} selected item${count !== 1 ? "s" : ""}? This cannot be undone.`;
  modal.classList.remove("hidden");

  const close = () => modal.classList.add("hidden");

  confirmBtn.onclick = async () => {
    confirmBtn.disabled = true;
    confirmBtn.textContent = "Deleting…";
    try {
      for (const name of _selection) {
        if (name.endsWith("/")) {
          await deleteFolderContents(name);
        } else {
          await deleteBlob(name);
        }
      }
      for (const delName of _selection) _audit("delete", delName);
      _selection.clear();
      _updateSelectionBar();
      close();
      _invalidateTreeChildren(_currentPrefix);
      _loadFiles(_currentPrefix);
      _showToast(`🗑️ ${count} item${count !== 1 ? "s" : ""} deleted`);
    } catch (err) {
      msgEl.innerHTML = `<span style="color:var(--error)">${_esc(err.message)}</span>`;
    } finally {
      confirmBtn.disabled = false;
      confirmBtn.textContent = "🗑️ Delete";
    }
  };
  cancelBtn.onclick = close;
  closeBtn.onclick  = close;
  modal.onclick = (e) => { if (e.target === modal) close(); };
}

// ── Delete modal ─────────────────────────────────────────────────────

function _showDeleteModal(item, type) {
  const modal      = _el("deleteModal");
  const msgEl      = _el("deleteModalMsg");
  const closeBtn   = _el("deleteModalClose");
  const cancelBtn  = _el("deleteCancelBtn");
  const confirmBtn = _el("deleteConfirmBtn");

  const label = type === "folder"
    ? `folder “${item.displayName}” and all its contents`
    : `file “${item.displayName}”`;
  msgEl.textContent = `Are you sure you want to permanently delete the ${label}? This cannot be undone.`;
  modal.classList.remove("hidden");

  const close = () => modal.classList.add("hidden");

  const doDelete = async () => {
    confirmBtn.disabled = true;
    confirmBtn.textContent = "Deleting…";
    try {
      if (type === "folder") {
        await deleteFolderContents(item.name);
      } else {
        await deleteBlob(item.name);
      }
      _audit("delete", item.name, { type });
      close();
      _invalidateTreeChildren(_currentPrefix);
      _loadFiles(_currentPrefix);
      _showToast(`🗑️ “${item.displayName}” deleted`);
    } catch (err) {
      msgEl.innerHTML = `<span style="color:var(--error)">${_esc(err.message)}</span>`;
    } finally {
      confirmBtn.disabled = false;
      confirmBtn.textContent = "🗑️ Delete";
    }
  };

  confirmBtn.onclick = doDelete;
  cancelBtn.onclick  = close;
  closeBtn.onclick   = close;
  modal.onclick = (e) => { if (e.target === modal) close(); };
}

// ── New item (folder / file) ───────────────────────────────

function _showNewModal() {
  const modal       = _el("newModal");
  const nameInput   = _el("newNameInput");
  const errEl       = _el("newModalError");
  const confirmBtn  = _el("newConfirmBtn");
  const closeBtn    = _el("newModalClose");
  const cancelBtn   = _el("newCancelBtn");
  const folderBtn   = _el("newTypeFolderBtn");
  const fileBtn     = _el("newTypeFileBtn");
  const contentWrap = _el("newFileContentWrap");
  const nameLabel   = _el("newNameLabel");

  let currentType = "folder";

  function setType(type) {
    currentType = type;
    folderBtn.classList.toggle("active", type === "folder");
    fileBtn.classList.toggle("active", type === "file");
    nameLabel.textContent    = type === "folder" ? "Folder name" : "File name";
    nameInput.placeholder    = type === "folder" ? "my-folder"   : "my-file.txt";
    contentWrap.classList.toggle("hidden", type === "folder");
  }

  setType("folder");
  nameInput.value = "";
  _el("newFileContent").value = "";
  errEl.textContent = "";
  errEl.classList.add("hidden");
  modal.classList.remove("hidden");
  setTimeout(() => nameInput.focus(), 50);

  folderBtn.onclick = () => setType("folder");
  fileBtn.onclick   = () => setType("file");

  const close = () => modal.classList.add("hidden");

  const doCreate = async () => {
    const name = nameInput.value.trim();
    if (!name) {
      errEl.textContent = "Please enter a name.";
      errEl.classList.remove("hidden");
      return;
    }
    if (name.includes("/") || name.includes("\\")) {
      errEl.textContent = 'Name cannot contain / or \\';
      errEl.classList.remove("hidden");
      return;
    }

    confirmBtn.disabled = true;
    confirmBtn.textContent = "Creating…";
    errEl.classList.add("hidden");

    try {
      if (currentType === "folder") {
        // Materialise the virtual folder with a hidden placeholder blob
        const placeholderPath = _currentPrefix + name + "/.keep";
        const file = new File([""], ".keep", { type: "application/octet-stream" });
        const u = getUser();
        const meta = {};
        if (u?.username) meta.uploaded_by_upn = u.username;
        if (u?.oid)      meta.uploaded_by_oid = u.oid;
        await uploadBlob(placeholderPath, file, null, meta);
        _audit("create", _currentPrefix + name + "/", { type: "folder" });
        close();
        _invalidateTreeChildren(_currentPrefix);
        _loadFiles(_currentPrefix + name + "/");
        _showToast(`\uD83D\uDCC1 Folder \u201C${name}\u201D created`);
      } else {
        const blobPath = _currentPrefix + name;
        const exists   = await blobExists(blobPath);
        if (exists) {
          errEl.textContent = `\u201C${name}\u201D already exists.`;
          errEl.classList.remove("hidden");
          confirmBtn.disabled = false;
          confirmBtn.textContent = "Create";
          return;
        }
        const content  = _el("newFileContent").value;
        const mimeType = _guessMimeType(name);
        const file     = new File([content], name, { type: mimeType });
        const u2 = getUser();
        const meta2 = {};
        if (u2?.username) meta2.uploaded_by_upn = u2.username;
        if (u2?.oid)      meta2.uploaded_by_oid = u2.oid;
        await uploadBlob(blobPath, file, null, meta2);
        _audit("create", blobPath, { type: "file" });
        close();
        _loadFiles(_currentPrefix);
        _showToast(`\uD83D\uDCC4 \u201C${name}\u201D created`);
      }
    } catch (err) {
      errEl.textContent = err.message;
      errEl.classList.remove("hidden");
    } finally {
      confirmBtn.disabled = false;
      confirmBtn.textContent = "Create";
    }
  };

  confirmBtn.onclick = doCreate;
  cancelBtn.onclick  = close;
  closeBtn.onclick   = close;
  modal.onclick = (e) => { if (e.target === modal) close(); };
  nameInput.onkeydown = (e) => {
    if (e.key === "Enter")  doCreate();
    if (e.key === "Escape") close();
  };
}

function _guessMimeType(filename) {
  const ext = filename.split(".").pop().toLowerCase();
  const map = {
    txt:  "text/plain",    md:   "text/markdown",    html: "text/html",
    htm:  "text/html",     css:  "text/css",          js:   "application/javascript",
    json: "application/json", xml: "application/xml", csv:  "text/csv",
    yaml: "text/yaml",     yml:  "text/yaml",          sh:   "text/plain",
    sql:  "text/plain",    log:  "text/plain",
  };
  return map[ext] || "application/octet-stream";
}

// ── Help modal ───────────────────────────────────────────────

function _showHelpModal() {
  const modal    = _el("helpModal");
  const closeBtn = _el("helpModalClose");
  const closeBtnFooter = _el("helpCloseBtn");

  modal.classList.remove("hidden");

  const close = () => modal.classList.add("hidden");
  closeBtn.onclick = close;
  closeBtnFooter.onclick = close;
  modal.addEventListener("click", (e) => { if (e.target === modal) close(); }, { once: true });
  const onKey = (e) => { if (e.key === "Escape") { close(); document.removeEventListener("keydown", onKey); } };
  document.addEventListener("keydown", onKey);
}

// ── Location info ──────────────────────────────────────────

async function _showInfoModal() {
  const modal    = _el("infoModal");
  const table    = _el("infoTable");
  const closeBtn = _el("infoModalClose");
  const closeBtnFooter = _el("infoCloseBtn");

  const close = () => modal.classList.add("hidden");
  closeBtn.onclick       = close;
  closeBtnFooter.onclick = close;
  modal.onclick = (e) => { if (e.target === modal) close(); };

  // Reset
  table.innerHTML = "";
  modal.classList.remove("hidden");

  const prefix = _currentPrefix;
  const displayPath = prefix ? prefix.replace(/\/$/, "") : "/";

  const addRow = (label, value, isHtml = false) => {
    const tr = document.createElement("tr");
    const tdLabel = document.createElement("td"); tdLabel.textContent = label;
    const tdVal   = document.createElement("td");
    if (isHtml) { tdVal.innerHTML = value; } else { tdVal.textContent = value; }
    tr.appendChild(tdLabel); tr.appendChild(tdVal);
    table.appendChild(tr);
  };

  addRow("Account",   CONFIG.storage.accountName);
  addRow("Container", CONFIG.storage.containerName);
  addRow("Path",      `<code style="font-size:12px">${_esc(displayPath)}</code>`, true);

  // Spinner row
  const spinnerTr = document.createElement("tr");
  spinnerTr.innerHTML = `<td colspan="2" style="color:var(--text-muted);font-size:12px;padding-top:10px">⏳ Scanning all levels…</td>`;
  table.appendChild(spinnerTr);

  try {
    const { totalFolders, totalFiles, totalSize } = await getFolderStats(prefix);
    spinnerTr.remove();
    addRow("Subfolders", totalFolders.toString());
    addRow("Files",     totalFiles.toString());
    addRow("Total size", formatFileSize(totalSize));
  } catch (e) {
    spinnerTr.remove();
    addRow("Error", e.message);
  }
}

let _toastTimer  = null;
let _toastMsgEl  = null;

function _updateToastMsg(msg) {
  if (_toastMsgEl) _toastMsgEl.textContent = msg;
}

function _showToast(msg, durationMs = 3000, actionLabel = null, actionFn = null) {
  const toast = _el("toast");
  toast.innerHTML = "";
  const msgSpan = document.createElement("span");
  msgSpan.textContent = msg;
  _toastMsgEl = msgSpan;
  toast.appendChild(msgSpan);
  if (actionLabel && actionFn) {
    const actionBtn = document.createElement("button");
    actionBtn.className = "toast-action";
    actionBtn.textContent = actionLabel;
    actionBtn.onclick = (e) => {
      e.stopPropagation();
      if (_toastTimer) clearTimeout(_toastTimer);
      toast.style.opacity = "0";
      setTimeout(() => toast.classList.add("hidden"), 300);
      actionFn();
    };
    toast.appendChild(actionBtn);
    toast.style.pointerEvents = "auto";
  } else {
    toast.style.pointerEvents = "none";
  }
  toast.classList.remove("hidden");
  toast.style.opacity = "1";
  if (_toastTimer) clearTimeout(_toastTimer);
  _toastTimer = setTimeout(() => {
    toast.style.opacity = "0";
    setTimeout(() => toast.classList.add("hidden"), 300);
  }, durationMs);
}
function _toggleUploadPanel() {
  const panel  = _el("uploadPanel");
  const isOpen = !panel.classList.contains("hidden");
  panel.classList.toggle("hidden", isOpen);
  _el("uploadBtn").classList.toggle("active", !isOpen);
}

// ── Upload queue ─────────────────────────────────────────────

let _uploadQueue   = [];
let _uploadCounter = 0;

function _queueFiles(fileList) {
  const maxBytes = (CONFIG.upload?.maxFileSizeMB || 0) * 1024 * 1024;

  Array.from(fileList).forEach((file) => {
    // Preserve folder structure for webkitdirectory / drag-dropped folders
    const relativePath = file.webkitRelativePath
      ? _currentPrefix + file.webkitRelativePath
      : _currentPrefix + file.name;

    const item = {
      id:       ++_uploadCounter,
      file,
      blobPath: relativePath,
      status:   "queued",
      progress: 0,
      error:    null,
    };

    if (maxBytes > 0 && file.size > maxBytes) {
      item.status = "error";
      item.error  = `Exceeds ${CONFIG.upload.maxFileSizeMB} MB limit`;
    }

    _uploadQueue.push(item);
    _renderUploadItem(item);
  });

  _el("uploadQueue").classList.remove("hidden");
  _el("uploadPanel").classList.remove("hidden");
  _el("uploadBtn").classList.add("active");
  _updateQueueCount();
  _startUpload();
}

async function _startUpload() {
  const pending = _uploadQueue.filter(i => i.status === "queued");
  if (pending.length === 0) return;

  // Check for conflicts in parallel and store the result on each item
  // so _processQueue can tell new files from overwrites when building metadata.
  const existsResults = await Promise.all(
    pending.map(item => blobExists(item.blobPath).catch(() => false))
  );
  pending.forEach((item, i) => { item.existed = existsResults[i]; });

  const conflicts = existsResults.filter(Boolean).length;
  if (conflicts === 0) {
    _processQueue(true);
  } else {
    _showOverwriteModal(conflicts, pending.length);
  }
}

function _showOverwriteModal(conflictCount, totalCount) {
  const modal      = _el("overwriteModal");
  const msgEl      = _el("overwriteModalMsg");
  const closeBtn   = _el("overwriteModalClose");
  const cancelBtn  = _el("overwriteCancelBtn");
  const skipBtn    = _el("overwriteSkipBtn");
  const confirmBtn = _el("overwriteConfirmBtn");

  const fileWord = (n) => `${n} file${n !== 1 ? "s" : ""}`;
  msgEl.textContent = `${fileWord(conflictCount)} of ${totalCount} selected already exist${conflictCount !== 1 ? "" : "s"} in this location. What would you like to do?`;
  modal.classList.remove("hidden");

  const close = () => modal.classList.add("hidden");

  const cancel = () => {
    close();
    // Remove all still-queued items from the queue and the UI list
    const cancelledIds = new Set(
      _uploadQueue.filter(i => i.status === "queued").map(i => i.id)
    );
    _uploadQueue = _uploadQueue.filter(i => i.status !== "queued");
    _el("uploadQueueList").querySelectorAll(".upload-item").forEach(el => {
      const id = parseInt(el.id.replace("upload-item-", ""), 10);
      if (cancelledIds.has(id)) el.remove();
    });
    _updateQueueCount();
  };

  confirmBtn.onclick = () => { close(); _processQueue(true);  };
  skipBtn.onclick    = () => { close(); _processQueue(false); };
  cancelBtn.onclick  = cancel;
  closeBtn.onclick   = cancel;
  modal.onclick = (e) => { if (e.target === modal) cancel(); };
}

async function _processQueue(overwrite) {
  const pending = _uploadQueue.filter(i => i.status === "queued");
  const user    = getUser();
  for (const item of pending) {
    item.status = "uploading";
    _updateItemUI(item);
    try {
      if (!overwrite && item.existed) {
        item.status = "skipped";
        item.error  = "File already exists (overwrite is off)";
      } else {
        const isOverwrite = overwrite && !!item.existed;
        const metadata    = await _buildUploadMetadata(isOverwrite, item.blobPath, user);
        await uploadBlob(item.blobPath, item.file, (pct) => {
          item.progress = pct;
          _updateItemUI(item);
        }, metadata);
        item.status   = "done";
        item.progress = 100;
        _audit("upload", item.blobPath, { size: item.file.size, overwrite: isOverwrite });
      }
    } catch (err) {
      item.status = "error";
      item.error  = err.message;
      console.error("[upload]", err.message);
    }
    _updateItemUI(item);
  }
  _updateQueueCount();
  // Refresh the file list once all pending uploads finish
  if (pending.length > 0) {
    // Invalidate every tree node whose children may have changed (folder uploads
    // can create new sub-prefixes at any depth).
    const dirtyPrefixes = new Set();
    for (const item of pending) {
      if (item.status !== "done") continue;
      // Walk every ancestor prefix of the uploaded blob
      const parts = item.blobPath.split("/");
      let acc = "";
      for (let i = 0; i < parts.length - 1; i++) {   // skip the filename itself
        acc += parts[i] + "/";
        dirtyPrefixes.add(acc);
      }
    }
    dirtyPrefixes.add(_currentPrefix);  // always refresh current level
    for (const pfx of dirtyPrefixes) _invalidateTreeChildren(pfx);

    _loadFiles(_currentPrefix);
  }
}

/**
 * Build the metadata object to attach to an uploaded blob.
 * - New file:    sets uploaded_by_upn / uploaded_by_oid.
 * - Overwrite:   reads existing metadata to preserve the original uploader,
 *               then adds / replaces last_edited_by_upn / last_edited_by_oid.
 */
async function _buildUploadMetadata(isOverwrite, blobPath, user) {
  const upn = user?.username || "";
  const oid = user?.oid      || "";
  if (!isOverwrite) {
    // Brand-new file
    const meta = {};
    if (upn) meta.uploaded_by_upn = upn;
    if (oid) meta.uploaded_by_oid = oid;
    return meta;
  }
  // Overwrite — preserve original uploader, stamp last-edited
  let meta = {};
  try { meta = await getBlobMetadata(blobPath); } catch (e) { console.warn("[upload] Could not fetch metadata:", e.message); }
  if (upn) meta.last_edited_by_upn = upn;
  if (oid) meta.last_edited_by_oid = oid;
  return meta;
}

function _renderUploadItem(item) {
  const list = _el("uploadQueueList");
  const div  = document.createElement("div");
  div.id        = `upload-item-${item.id}`;
  div.className = "upload-item";
  div.innerHTML = `
    <span class="upload-item-icon">${_esc(getFileIcon(item.file.name))}</span>
    <div class="upload-item-info">
      <span class="upload-item-name" title="${_esc(item.blobPath)}">${_esc(item.file.name)}</span>
      <span class="upload-item-size">${formatFileSize(item.file.size)}</span>
      <span class="upload-item-error hidden"></span>
    </div>
    <div class="upload-item-right">
      <div class="progress-track"><div class="progress-fill" style="width:0%"></div></div>
      <span class="upload-item-status">Queued</span>
    </div>`;
  if (item.status === "error") {
    div.querySelector(".upload-item-status").textContent = "✗ Error";
    div.querySelector(".upload-item-status").className   = "upload-item-status status-error";
    const errSpan = div.querySelector(".upload-item-error");
    errSpan.textContent = item.error || "Unknown error";
    errSpan.classList.remove("hidden");
  }
  list.appendChild(div);
}

function _updateItemUI(item) {
  const div = _el(`upload-item-${item.id}`);
  if (!div) return;
  const fill   = div.querySelector(".progress-fill");
  const status = div.querySelector(".upload-item-status");
  fill.style.width = `${item.progress}%`;
  switch (item.status) {
    case "done":
      status.textContent = "✓ Done";
      status.className   = "upload-item-status status-done";
      fill.style.width   = "100%";
      break;
    case "error": {
      status.textContent = "✗ Error";
      status.className   = "upload-item-status status-error";
      const errSpan = div.querySelector(".upload-item-error");
      if (errSpan) {
        errSpan.textContent = item.error || "Unknown error";
        errSpan.classList.remove("hidden");
      }
      break;
    }
    case "uploading":
      status.textContent = `${item.progress}%`;
      status.className   = "upload-item-status status-uploading";
      break;
    case "skipped": {
      status.textContent = "⏭ Skipped";
      status.className   = "upload-item-status status-skipped";
      const skipSpan = div.querySelector(".upload-item-error");
      if (skipSpan) {
        skipSpan.textContent = item.error || "Skipped";
        skipSpan.classList.remove("hidden");
      }
      break;
    }
    default:
      status.textContent = "Queued";
      status.className   = "upload-item-status";
  }
}

function _clearCompleted() {
  _uploadQueue = _uploadQueue.filter(i => i.status !== "done" && i.status !== "skipped");
  [..._el("uploadQueueList").children].forEach(el => {
    const id = parseInt(el.id.replace("upload-item-", ""), 10);
    if (!_uploadQueue.find(i => i.id === id)) el.remove();
  });
  _updateQueueCount();
  // Hide the queue panel when nothing remains
  if (_uploadQueue.length === 0) {
    _el("uploadQueue").classList.add("hidden");
  }
}

function _updateQueueCount() {
  const done  = _uploadQueue.filter(i => i.status === "done").length;
  const total = _uploadQueue.length;
  _el("uploadQueueCount").textContent = total ? `${done} of ${total} uploaded` : "";
}

function _initDragDrop() {
  const area = _el("contentArea");

  area.addEventListener("dragover", (e) => {
    e.preventDefault();
    area.classList.add("drag-over");
    // Auto-open the upload panel when dragging files over the app
    _el("uploadPanel").classList.remove("hidden");
    _el("uploadBtn").classList.add("active");
  });

  area.addEventListener("dragleave", (e) => {
    if (!area.contains(e.relatedTarget)) {
      area.classList.remove("drag-over");
    }
  });

  area.addEventListener("drop", (e) => {
    e.preventDefault();
    area.classList.remove("drag-over");
    if (e.dataTransfer.files.length > 0) {
      _queueFiles(e.dataTransfer.files);
    }
  });
}

// ── Folder tree ──────────────────────────────────────────────

let _treeVisible = true;

function _initFolderTree() {
  const treeRoot = _el("treeRoot");
  treeRoot.innerHTML = "";

  // Root node — always visible and starts expanded
  const rootNode = _makeTreeNode(CONFIG.storage.containerName, "");
  rootNode.classList.add("expanded");
  const rootChildren = rootNode.querySelector(".tree-children");
  rootChildren.dataset.loaded = "1";
  treeRoot.appendChild(rootNode);
  _loadTreeChildren("", rootNode);

  // Wire panel-toggle button
  _el("treeToggleBtn").onclick = _toggleFolderTree;

  // Wire expand-all button — recursively loads and expands every node
  _el("treeExpandAllBtn").onclick = () => _expandAllTreeNodes(_el("treeRoot"));

  // Wire collapse-all button
  _el("treeCollapseAllBtn").onclick = () => {
    _el("treeRoot").querySelectorAll(".tree-node.expanded").forEach(n => {
      if (n.dataset.prefix !== "") n.classList.remove("expanded");
    });
  };

  // Default visibility: show on wide screens, hide on tablet/phone
  _treeVisible = window.innerWidth > 900;
  _el("folderTree").classList.toggle("tree-hidden", !_treeVisible);
  _el("treeToggleBtn").classList.toggle("active", _treeVisible);
  _updateTreeToggleArrow();

  // Initialise right-click context menu on tree nodes
  _initTreeContextMenu();
}

function _toggleFolderTree() {
  _treeVisible = !_treeVisible;
  _el("folderTree").classList.toggle("tree-hidden", !_treeVisible);
  _el("treeToggleBtn").classList.toggle("active", _treeVisible);
  _updateTreeToggleArrow();
}

function _updateTreeToggleArrow() {
  const btn = _el("treeToggleBtn");
  // ◀ when tree is open (click to close), ▶ when tree is hidden (click to open)
  btn.innerHTML = (_treeVisible ? "\u25C4" : "\u25BA") + " <span class=\"btn-label\">Tree</span>";
}

/**
 * Recursively expand every folder in the tree, loading children on demand.
 * Awaits each level before descending so the DOM is populated before recursing.
 */
async function _expandAllTreeNodes(container) {
  const nodes = container.querySelectorAll(":scope > .tree-node");
  for (const node of nodes) {
    const childrenEl = node.querySelector(":scope > .tree-children");
    if (!childrenEl) continue;
    // Load children from the API if not yet fetched
    if (!childrenEl.dataset.loaded) {
      childrenEl.dataset.loaded = "1";
      await _loadTreeChildren(node.dataset.prefix, node);
    }
    node.classList.add("expanded");
    // Recurse into this node's children container
    await _expandAllTreeNodes(childrenEl);
  }
}

function _makeTreeNode(displayName, prefix) {
  const node = document.createElement("div");
  node.className = "tree-node";
  node.dataset.prefix = prefix;

  const row = document.createElement("div");
  row.className = "tree-node-row";

  const chevron = document.createElement("button");
  chevron.type      = "button";
  chevron.className = "tree-chevron";
  chevron.innerHTML = "&#9654;"; // ▶
  chevron.title     = "Expand / collapse";

  const icon = document.createElement("span");
  icon.className  = "tree-icon";
  icon.textContent = prefix === "" ? "\uD83D\uDCE6" : "\uD83D\uDCC1";

  const label = document.createElement("a");
  label.className   = "tree-label";
  label.href        = "#";
  label.textContent = displayName;
  label.title       = prefix || "Root";

  row.appendChild(chevron);
  row.appendChild(icon);
  row.appendChild(label);

  const childrenEl = document.createElement("div");
  childrenEl.className = "tree-children";

  node.appendChild(row);
  node.appendChild(childrenEl);

  function ensureChildrenLoaded() {
    if (!childrenEl.dataset.loaded) {
      childrenEl.dataset.loaded = "1";
      _loadTreeChildren(prefix, node);
    }
  }

  chevron.addEventListener("click", (e) => {
    e.preventDefault();
    e.stopPropagation();
    const isNowExpanded = node.classList.toggle("expanded");
    if (isNowExpanded) ensureChildrenLoaded();
  });

  label.addEventListener("click", (e) => {
    e.preventDefault();
    if (!node.classList.contains("expanded")) {
      node.classList.add("expanded");
      ensureChildrenLoaded();
    }
    _loadFiles(prefix);
  });

  // Right-click context menu on the tree row
  row.addEventListener("contextmenu", (e) => {
    e.preventDefault();
    e.stopPropagation();
    _showTreeContextMenu(e, prefix);
  });

  return node;
}

async function _loadTreeChildren(prefix, parentNode) {
  const childrenEl = parentNode.querySelector(":scope > .tree-children");
  childrenEl.innerHTML = `<div class="tree-loading">Loading\u2026</div>`;

  try {
    let { folders } = await listBlobsAtPrefix(prefix);
    // Hide audit folder from the tree
    if (prefix === "") folders = folders.filter((f) => f.displayName !== _AUDIT_FOLDER);
    childrenEl.innerHTML = "";

    if (folders.length === 0) {
      // Leaf node — hide the expand chevron
      const ch = parentNode.querySelector(":scope > .tree-node-row > .tree-chevron");
      if (ch) ch.style.visibility = "hidden";
      return;
    }

    for (const folder of folders) {
      childrenEl.appendChild(_makeTreeNode(folder.displayName, folder.name));
    }
  } catch (err) {
    // In SAS mode a listing 403 is expected when the token has no list permission —
    // treat the node as a leaf rather than showing an error.
    if (isSasMode() && /403|Access denied/i.test(err?.message || "")) {
      const ch = parentNode.querySelector(":scope > .tree-node-row > .tree-chevron");
      if (ch) ch.style.visibility = "hidden";
      childrenEl.innerHTML = "";
      return;
    }
    childrenEl.innerHTML = `<div class="tree-loading tree-load-err">Failed to load</div>`;
  }
}

// ── Tree context menu ────────────────────────────────────────

function _showTreeContextMenu(event, prefix) {
  const menu = _el("treeContextMenu");
  if (!menu) return;

  // Show / hide permission-gated items
  menu.querySelector('[data-action="new"]').classList.toggle("hidden", !_canUpload);
  menu.querySelector('[data-action="upload"]').classList.toggle("hidden", !_canUpload);
  menu.querySelector('[data-action="copy"]').classList.toggle("hidden", !_canCopyItems());
  menu.querySelector('[data-action="move"]').classList.toggle("hidden", !_canMoveItems());
  menu.querySelector('[data-action="rename"]').classList.toggle("hidden", !_canRenameItems() || prefix === "");
  menu.querySelector('[data-action="delete"]').classList.toggle("hidden", !_canDeleteItems() || prefix === "");
  menu.querySelector('[data-action="sas"]').classList.toggle("hidden", !_canSas());

  // Hide email sub-items when not available
  const canEmail = _canEmail();
  menu.querySelectorAll('.ctx-email-item').forEach(el => el.classList.toggle("hidden", !canEmail));

  // Hide danger separator if delete is hidden
  const dangerSep = menu.querySelector('.ctx-sep-danger');
  if (dangerSep) dangerSep.classList.toggle("hidden", !_canDeleteItems() || prefix === "");

  // Position the menu at the cursor, clamped to the viewport
  menu.classList.remove("hidden");
  const menuW = menu.offsetWidth;
  const menuH = menu.offsetHeight;
  let x = event.clientX;
  let y = event.clientY;
  if (x + menuW > window.innerWidth)  x = window.innerWidth  - menuW - 4;
  if (y + menuH > window.innerHeight) y = window.innerHeight - menuH - 4;
  if (x < 0) x = 4;
  if (y < 0) y = 4;
  menu.style.left = x + "px";
  menu.style.top  = y + "px";

  // Flip submenu left if it would overflow the right edge
  const submenu = menu.querySelector('.ctx-submenu');
  if (submenu) {
    submenu.classList.remove('flip-left');
    const menuRight = x + menu.offsetWidth;
    if (menuRight + 180 > window.innerWidth) submenu.classList.add('flip-left');
  }

  // Store which prefix this menu targets
  menu.dataset.prefix = prefix;
}

function _hideTreeContextMenu() {
  const menu = _el("treeContextMenu");
  if (menu) menu.classList.add("hidden");
}

let _treeContextMenuInitialized = false;

function _initTreeContextMenu() {
  const menu = _el("treeContextMenu");
  if (!menu) return;

  // Register document/window listeners only once to avoid duplicate handlers
  if (!_treeContextMenuInitialized) {
    _treeContextMenuInitialized = true;

    // Close on any outside click or Escape
    document.addEventListener("click",   _hideTreeContextMenu);
    document.addEventListener("contextmenu", () => {
      // Defer so the tree-node handler can re-open if needed
      setTimeout(_hideTreeContextMenu, 0);
    });
    document.addEventListener("keydown", (e) => {
      if (e.key === "Escape") _hideTreeContextMenu();
    });
    window.addEventListener("blur", _hideTreeContextMenu);
    window.addEventListener("scroll", _hideTreeContextMenu, true);

    // Action dispatcher
    menu.addEventListener("click", (e) => {
      const item = e.target.closest(".tree-context-menu-item");
      if (!item) return;
      const action = item.dataset.action;
      // Items without data-action (e.g. submenu triggers) are non-actions; leave menu open
      if (!action) return;
      const prefix = menu.dataset.prefix || "";
      _hideTreeContextMenu();

      // Navigate to the folder first so the actions target the right prefix
      if (_currentPrefix !== prefix) {
        _loadFiles(prefix).then(() => _runTreeContextAction(action, prefix)).catch(() => {});
      } else {
        _runTreeContextAction(action, prefix);
      }
    });
  }
}

function _treeFolderItem(prefix) {
  // Build a minimal folder item object matching the shape expected by the action modals
  const displayName = prefix ? prefix.replace(/\/$/, "").split("/").pop() : CONFIG.storage.containerName;
  return { name: prefix, displayName };
}

function _runTreeContextAction(action, prefix) {
  const item = _treeFolderItem(prefix);
  switch (action) {
    case "new":       _showNewModal();                        break;
    case "upload":    _toggleUploadPanel();                   break;
    case "download":  _downloadCurrentLevel();                break;
    case "info":      _showInfoModal();                       break;
    case "refresh":   _loadFiles(_currentPrefix);             break;
    case "report":    _exportReport();                        break;
    case "copy-blob-url": _treeCopyBlobUrl(item);              break;
    case "copy-app-link": _treeCopyAppLink(item);              break;
    case "email-blob-url": _treeEmailBlobUrl(item);            break;
    case "email-app-link": _treeEmailAppLink(item);            break;
    case "copy":      _showCopyModal(item, "folder");         break;
    case "move":      _showMoveModal(item, "folder");         break;
    case "rename":    _showRenameModal(item, "folder");       break;
    case "delete":    _showDeleteModal(item, "folder");       break;
    case "sas":       _showSasModal(item, true);              break;
  }
}

function _treeBuildUrls(item) {
  const { accountName, containerName } = CONFIG.storage;
  const encoded = item.name.split("/").map(encodeURIComponent).join("/");
  const blobUrl = `https://${accountName}.blob.core.windows.net/${containerName}/${encoded}`;
  const appUrl  = _buildAppUrl(item.name);
  return { blobUrl, appUrl };
}

function _treeCopyBlobUrl(item) {
  const { blobUrl } = _treeBuildUrls(item);
  _copyToClipboard(blobUrl);
}

function _treeCopyAppLink(item) {
  const { appUrl } = _treeBuildUrls(item);
  if (CONFIG.app.allowDownload) {
    const dlFn = () => _handleFolderDownload(item);
    const doToast = () => _showToast("\uD83D\uDCCB App link copied!", 8000, "\u2B07 Download", dlFn);
    if (navigator.clipboard && window.isSecureContext) {
      navigator.clipboard.writeText(appUrl).then(doToast).catch(() => { _fallbackCopy(appUrl); doToast(); });
    } else {
      _fallbackCopy(appUrl);
      doToast();
    }
  } else {
    _copyToClipboard(appUrl);
  }
}

function _treeEmailBlobUrl(item) {
  const { blobUrl } = _treeBuildUrls(item);
  _showEmailComposeModal(item.displayName, blobUrl);
}

function _treeEmailAppLink(item) {
  const { appUrl } = _treeBuildUrls(item);
  _showEmailComposeModal(item.displayName, appUrl);
}

/**
 * Remove the "loaded" cache flag from a tree node's children container
 * so the next _syncTreeToPrefix walk will re-fetch its children from the API.
 */
function _invalidateTreeChildren(prefix) {
  const treeRoot = _el("treeRoot");
  if (!treeRoot) return;
  for (const n of treeRoot.querySelectorAll(".tree-node")) {
    if (n.dataset.prefix === prefix) {
      const ch = n.querySelector(":scope > .tree-children");
      if (ch) { delete ch.dataset.loaded; ch.innerHTML = ""; }
      // Restore chevron visibility in case it was hidden (leaf → parent)
      const chevron = n.querySelector(":scope > .tree-node-row > .tree-chevron");
      if (chevron) chevron.style.visibility = "";
      break;
    }
  }
}

async function _syncTreeToPrefix(prefix) {
  const treeRoot = _el("treeRoot");
  if (!treeRoot) return;

  // Remove all active highlights
  treeRoot.querySelectorAll(".tree-node-row.active")
    .forEach(r => r.classList.remove("active"));

  // The tree has one root node (prefix = ""). All folder nodes live inside it.
  const rootNode = _findTreeChild(treeRoot, "");
  if (!rootNode) return;

  if (!prefix) {
    // Viewing the container root
    rootNode.querySelector(":scope > .tree-node-row").classList.add("active");
    return;
  }

  // Ensure root is expanded and its children are loaded
  if (!rootNode.classList.contains("expanded")) rootNode.classList.add("expanded");
  const rootChildrenEl = rootNode.querySelector(":scope > .tree-children");
  if (!rootChildrenEl.dataset.loaded) {
    rootChildrenEl.dataset.loaded = "1";
    await _loadTreeChildren("", rootNode);
  }

  // Walk path segments, searching inside each level's children container
  const parts = prefix.split("/").filter(Boolean);
  let accumulated = "";
  let searchEl = rootChildrenEl; // start inside root's children, not treeRoot

  for (let i = 0; i < parts.length; i++) {
    accumulated += parts[i] + "/";
    const node = _findTreeChild(searchEl, accumulated);
    if (!node) break;

    // Expand every node along the active path
    if (!node.classList.contains("expanded")) node.classList.add("expanded");
    const childEl = node.querySelector(":scope > .tree-children");
    if (!childEl.dataset.loaded) {
      childEl.dataset.loaded = "1";
      await _loadTreeChildren(accumulated, node);
    }

    if (i === parts.length - 1) {
      // Highlight the target folder
      const row = node.querySelector(":scope > .tree-node-row");
      row.classList.add("active");
      row.scrollIntoView({ block: "nearest" });
    }

    searchEl = childEl;
  }
}

function _findTreeChild(containerEl, prefix) {
  for (const child of containerEl.children) {
    if (child.classList.contains("tree-node") && child.dataset.prefix === prefix) {
      return child;
    }
  }
  return null;
}

// ── UI helpers ───────────────────────────────────────────────

function _showLoading(visible) {
  _el("loadingOverlay").classList.toggle("hidden", !visible);
}

function _showError(msg) {
  const banner = _el("errorMessage");
  banner.textContent = msg;
  banner.classList.remove("hidden");
  _el("emptyState").classList.add("hidden");
}

function _hideError() {
  _el("errorMessage").classList.add("hidden");
  _el("networkErrorHelp").classList.add("hidden");
}

function _el(id) {
  return document.getElementById(id);
}

/** Reusable element for HTML-escaping (avoids creating a new element per call). */
const _escDiv = document.createElement("div");

function _esc(text) {
  _escDiv.textContent = text;
  return _escDiv.innerHTML;
}

// ── SAS generator ────────────────────────────────────────────

/**
 * Open the SAS generator modal for a file or folder.
 * @param {{ name: string, displayName: string }} item
 * @param {boolean} isFolder  true = container SAS, false = blob SAS
 */
function _showSasModal(item, isFolder) {
  const modal       = _el("sasModal");
  const titleEl     = _el("sasModalTitle");
  const infoEl      = _el("sasItemInfo");
  const startEl     = _el("sasStart");
  const expiryEl    = _el("sasExpiry");
  const ipEl        = _el("sasIp");
  const permGroup   = _el("sasPermissions");
  const folderNote  = _el("sasFolderNote");
  const errEl       = _el("sasError");
  const resultWrap  = _el("sasResultWrap");
  const resultEl    = _el("sasResult");
  const expiresNote = _el("sasExpiresNote");
  const generateBtn = _el("sasGenerateBtn");
  const cancelBtn   = _el("sasCancelBtn");
  const copyBtn     = _el("sasCopyBtn");
  const emailSasBtn = _el("sasEmailBtn");
  const closeBtn    = _el("sasModalClose");

  const close = () => modal.classList.add("hidden");
  closeBtn.onclick  = close;
  cancelBtn.onclick = close;
  modal.onclick = (e) => { if (e.target === modal) close(); };

  // Show/hide email button based on capability
  emailSasBtn.classList.toggle("hidden", !_canEmail());

  // ── Header & item info ──────────────────────────────────────
  titleEl.textContent = isFolder ? "Generate folder SAS" : "Generate file SAS";
  infoEl.textContent  = (isFolder ? "\uD83D\uDCC1 " : getFileIcon(item.displayName) + " ") + (item.displayName || item.name);

  folderNote.classList.toggle("hidden", !isFolder);

  // ── Default datetimes (start = blank, expiry = now + 1 h) ───
  const _fmtLocal = (d) => {
    const p = (n) => String(n).padStart(2, "0");
    return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}T${p(d.getHours())}:${p(d.getMinutes())}`;
  };
  startEl.value  = "";
  expiryEl.value = _fmtLocal(new Date(Date.now() + 60 * 60 * 1000));
  ipEl.value     = "";

  // ── Permission checkboxes ────────────────────────────────────
  const showPerms = CONFIG.app.sasShowPermissions !== false;
  permGroup.closest(".sas-row").classList.toggle("hidden", !showPerms);
  permGroup.innerHTML = "";

  if (showPerms) {
    const permDefs = isFolder
      ? [
          { char: "r", label: "Read",   checked: true  },
          { char: "l", label: "List",   checked: true  },
          { char: "w", label: "Write",  checked: false },
          { char: "d", label: "Delete", checked: false },
        ]
      : [
          { char: "r", label: "Read",   checked: true  },
          { char: "w", label: "Write",  checked: false },
          { char: "d", label: "Delete", checked: false },
        ];

    permDefs.forEach(({ char, label, checked }) => {
      const lbl = document.createElement("label");
      lbl.className = "sas-perm-label";
      const cb = document.createElement("input");
      cb.type    = "checkbox";
      cb.value   = char;
      cb.checked = checked;
      lbl.appendChild(cb);
      lbl.append(" " + label);
      permGroup.appendChild(lbl);
    });
  }

  // ── Reset result area ────────────────────────────────────────
  errEl.textContent = "";
  errEl.classList.add("hidden");
  resultWrap.classList.add("hidden");
  resultEl.value = "";

  modal.classList.remove("hidden");

  // ── Generate button ──────────────────────────────────────────
  generateBtn.onclick = async () => {
    errEl.classList.add("hidden");
    resultWrap.classList.add("hidden");

    if (!expiryEl.value) {
      errEl.textContent = "Please set an expiry date/time.";
      errEl.classList.remove("hidden");
      return;
    }
    const expiryDate = new Date(expiryEl.value);
    if (expiryDate <= new Date()) {
      errEl.textContent = "Expiry must be in the future.";
      errEl.classList.remove("hidden");
      return;
    }

    // Build permissions string in canonical order
    let sp;
    if (CONFIG.app.sasShowPermissions === false) {
      // Permissions hidden — always issue read-only (+ list for folders)
      sp = isFolder ? "rl" : "r";
    } else {
      const permOrder = isFolder ? "racwdxltfmeop" : "racwdxtlfsop";
      const selected  = new Set(
        [...permGroup.querySelectorAll("input[type=checkbox]:checked")].map((c) => c.value)
      );
      sp = permOrder.split("").filter((c) => selected.has(c)).join("");
      if (!sp) {
        errEl.textContent = "Select at least one permission.";
        errEl.classList.remove("hidden");
        return;
      }
    }

    // Convert datetime-local (local time) to UTC ISO 8601
    const toUtcIso = (val) =>
      val ? new Date(val).toISOString().replace(/\.\d{3}Z$/, "Z") : "";

    generateBtn.disabled    = true;
    generateBtn.textContent = "Generating\u2026";

    try {
      const sasUrl = await generateSasToken(
        isFolder ? "" : item.name,
        isFolder,
        {
          start:       toUtcIso(startEl.value),
          expiry:      toUtcIso(expiryEl.value),
          permissions: sp,
          ip:          ipEl.value.trim(),
        }
      );
      resultEl.value = sasUrl;
      _audit("sas", isFolder ? (item.name || "/") : item.name, {
        permissions: sp,
        expiry: toUtcIso(expiryEl.value),
        ip: ipEl.value.trim() || undefined,
      });
      resultWrap.classList.remove("hidden");
      expiresNote.textContent = `Expires ${new Date(expiryEl.value).toLocaleString()}`;
      resultEl.select();
    } catch (err) {
      errEl.textContent = err.message;
      errEl.classList.remove("hidden");
    } finally {
      generateBtn.disabled    = false;
      generateBtn.textContent = "\uD83D\uDD11 Generate";
    }
  };

  // ── Copy SAS URL button ──────────────────────────────────────
  copyBtn.onclick = async () => {
    if (!resultEl.value) return;
    try {
      await navigator.clipboard.writeText(resultEl.value);
      const orig = copyBtn.textContent;
      copyBtn.textContent = "\u2705 Copied";
      setTimeout(() => { copyBtn.textContent = orig; }, 2000);
    } catch {
      _fallbackCopy(resultEl.value);
    }
  };

  // ── Email SAS URL button ────────────────────────────────────
  emailSasBtn.onclick = () => {
    if (!resultEl.value) return;
    const itemName = item.displayName || item.name.split("/").filter(Boolean).pop() || item.name;
    close();
    _showEmailComposeModal(itemName, resultEl.value);
  };
}

// ── Theme toggle (dark / light mode) ─────────────────────────

const _THEME_STORAGE_KEY = "be_theme";

function _initTheme() {
  const saved = localStorage.getItem(_THEME_STORAGE_KEY);
  if (saved === "dark" || saved === "light") {
    document.documentElement.setAttribute("data-theme", saved);
  } else if (window.matchMedia("(prefers-color-scheme: dark)").matches) {
    document.documentElement.setAttribute("data-theme", "dark");
  }
  _updateThemeIcons();
  // Listen for OS-level theme changes when user has no explicit preference
  window.matchMedia("(prefers-color-scheme: dark)").addEventListener("change", (e) => {
    if (localStorage.getItem(_THEME_STORAGE_KEY)) return; // user chose manually
    document.documentElement.setAttribute("data-theme", e.matches ? "dark" : "light");
    _updateThemeIcons();
  });
}

function _toggleTheme() {
  const isDark = document.documentElement.getAttribute("data-theme") === "dark";
  const next = isDark ? "light" : "dark";
  document.documentElement.setAttribute("data-theme", next);
  localStorage.setItem(_THEME_STORAGE_KEY, next);
  _updateThemeIcons();
}

function _updateThemeIcons() {
  const isDark = document.documentElement.getAttribute("data-theme") === "dark";
  // Icon shows the action: ☀️ = "switch to light", 🌙 = "switch to dark"
  const icon = isDark ? "\u2600\ufe0f" : "\uD83C\uDF19";
  document.querySelectorAll(".theme-toggle").forEach((btn) => {
    // Update only the first text node so the <span class="btn-label"> child is preserved
    const textNode = Array.from(btn.childNodes).find((n) => n.nodeType === Node.TEXT_NODE);
    if (textNode) { textNode.textContent = icon + " "; } else { btn.prepend(icon + " "); }
  });
}

// Apply theme immediately (before DOMContentLoaded so no flash-of-wrong-theme)
_initTheme();

// Wire all theme toggle buttons once the DOM is ready
document.addEventListener("DOMContentLoaded", () => {
  document.querySelectorAll(".theme-toggle").forEach((btn) => {
    btn.addEventListener("click", _toggleTheme);
  });
});

// ── Keyboard shortcuts ───────────────────────────────────────

document.addEventListener("keydown", (e) => {
  // Skip when user is typing in an input, textarea, or contenteditable
  const tag = (e.target.tagName || "").toLowerCase();
  const isEditing = tag === "input" || tag === "textarea" || e.target.isContentEditable;

  // Ctrl/Cmd + K — toggle search bar
  if ((e.ctrlKey || e.metaKey) && e.key === "k") {
    e.preventDefault();
    if (!_el("mainApp")?.classList.contains("hidden")) _toggleSearchBar();
    return;
  }

  // Ctrl/Cmd + U — toggle upload panel
  if ((e.ctrlKey || e.metaKey) && e.key === "u") {
    e.preventDefault();
    if (!_el("mainApp")?.classList.contains("hidden") && _canUpload) {
      _el("uploadBtn")?.click();
    }
    return;
  }

  // F5 — refresh (prevent default browser reload and refresh the listing instead)
  if (e.key === "F5" && !e.ctrlKey && !e.metaKey && !e.shiftKey) {
    if (!_el("mainApp")?.classList.contains("hidden")) {
      e.preventDefault();
      _loadFiles(_currentPrefix);
      return;
    }
  }

  // Skip remaining shortcuts when user is editing text
  if (isEditing) return;

  // Backspace — navigate to parent folder
  if (e.key === "Backspace") {
    if (!_el("mainApp")?.classList.contains("hidden") && _currentPrefix) {
      e.preventDefault();
      _goUp();
    }
    return;
  }

  // ? — show help modal
  if (e.key === "?") {
    _showHelpModal();
    return;
  }
});