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
let _containerSearchResults = null;   // null = normal view; array = container-wide search hits
let _searchDebounceTimer    = null;

// sessionStorage key for persisting the user's storage selection
const _STORAGE_SELECTION_KEY = "be_storage_selection";

// Derived permission flags — depend on both RBAC probe and config switches
function _canRenameItems() { return _canUpload && (CONFIG.app.allowRename !== false); }
function _canDeleteItems() { return _canUpload && (CONFIG.app.allowDelete !== false); }
function _canSas()         { return _canUpload && CONFIG.app.allowSas !== false; }
function _canEditItems()   { return _canUpload; }

// ── Entry point ──────────────────────────────────────────────

document.addEventListener("DOMContentLoaded", async () => {
  _showLoading(true);
  try {
    const account = await initAuth();
    if (account) {
      // Restore a previously chosen storage selection (survives page refresh)
      _restoreStorageSelection();
      // Show picker when no storage is configured, unless the picker is disabled
      const pickerEnabled = CONFIG.app.allowStoragePicker !== false;
      if (pickerEnabled && (!CONFIG.storage.accountName || !CONFIG.storage.containerName)) {
        _showLoading(false);
        _showPickerPage(account);
      } else {
        _bootApp(account);
      }
    } else {
      // No stored session — redirect straight to Microsoft login.
      // If the user already has an active Entra ID / M365 session the
      // browser will be sent straight back without any login prompt (SSO).
      await signIn();
      // signIn() navigates away — nothing below executes.
    }
  } catch (err) {
    console.error("[app] Init error:", err);
    // Show the manual sign-in page only as a last-resort fallback
    // (e.g. the auth server returned an unexpected error).
    _showLoading(false);
    _showSignInPage();
  }
});

// ── Sign-in / sign-out ───────────────────────────────────────

function _showSignInPage() {
  _el("signInPage").classList.remove("hidden");
  _el("mainApp").classList.add("hidden");
  _el("pickerPage").classList.add("hidden");
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
}

// ── Storage picker ───────────────────────────────────────────

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
  document.title = "Select Storage — " + (CONFIG.app.title || "Blob Storage Explorer");

  // Populate user info in picker header
  const displayName = account.name || account.username;
  const upn          = account.username;
  _el("pickerUserName").textContent = upn ? `${displayName} (${upn})` : displayName;
  _el("pickerSignOutBtn").onclick = () => signOut();

  // Show the Back button only when the user is already inside a session
  // (i.e. they came here via the Change button, not on initial sign-in)
  const hasActiveSession = !!(CONFIG.storage.accountName && CONFIG.storage.containerName);
  const backBtn = _el("pickerBackBtn");
  backBtn.classList.toggle("hidden", !hasActiveSession);
  backBtn.onclick = hasActiveSession ? () => {
    _el("pickerPage").classList.add("hidden");
    _el("mainApp").classList.remove("hidden");
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

  let allAccountData = []; // flat array of { account, containers }

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

    // Fetch containers for all accounts in parallel — used to filter out empties
    const containerFetches = await Promise.allSettled(
      accountEntries.map(({ acct }) =>
        listContainers(acct.subscriptionId, acct.resourceGroup, acct.name)
          .then(containers => ({ name: acct.name, containers }))
      )
    );
    const containerMap = new Map();
    containerFetches.forEach(r => {
      if (r.status === "fulfilled") containerMap.set(r.value.name, r.value.containers);
    });

    // Group by subscription, skipping accounts that have zero containers
    const subGroups = new Map();
    for (const { sub, acct } of accountEntries) {
      const containers = containerMap.get(acct.name);
      if (containers !== undefined && containers.length === 0) continue;
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
  return raw ? JSON.parse(raw) : null;
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

  _el("signOutBtn").onclick     = () => signOut();
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
  const hashPath = window.location.hash.slice(1);
  _listingPromise = _loadFiles(hashPath || "");

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
    badge.textContent = hasPermission ? "\u270f\ufe0f Writer" : "\uD83D\uDCD6 Reader";
    badge.classList.remove("hidden");
    if (!hasPermission) return;
    _canUpload = true;
    _el("uploadBtn").classList.remove("hidden");
    _el("newBtn").classList.remove("hidden");
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
  _currentPrefix = prefix;
  history.replaceState(null, "", prefix ? `#${prefix}` : window.location.pathname);
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
    const { files: sortedFiles } = _sortItems([], _containerSearchResults);
    banner.classList.remove("hidden");
    banner.innerHTML =
      `\uD83D\uDD0D <strong>${sortedFiles.length}</strong> result${sortedFiles.length !== 1 ? "s" : ""} `
      + `for \u201c<strong>${_esc(term)}</strong>\u201d across the entire container`;
    countLabel.textContent = `${sortedFiles.length} result${sortedFiles.length !== 1 ? "s" : ""}`;

    if (sortedFiles.length === 0) {
      emptyState.classList.remove("hidden");
      emptyState.querySelector("p").textContent = `No files found matching \u201c${term}\u201d`;
      return;
    }
    emptyState.classList.add("hidden");
    if (_viewMode === "list") {
      _renderSearchResultsListView(container, sortedFiles);
    } else {
      _renderGridView(container, [], sortedFiles);
    }
    return;
  }

  // ── Normal folder view ──────────────────────────────────────
  banner.classList.add("hidden");
  const term = (_el("searchInput")?.value || "").trim().toLowerCase();
  let folders = _cachedFolders;
  let files   = _cachedFiles;

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
    // Whole-container: debounce 500 ms before hitting the API
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
        <button class="btn-action btn-props" title="Properties">ℹ️ Info</button>
        <button class="btn-action btn-copy-url" title="Copy URL">🔗 Copy URL</button>
        ${_canSas() ? `<button class="btn-action btn-sas" title="Generate SAS">🔑 SAS</button>` : ""}
        ${CONFIG.app.allowDownload ? `<button class="btn-action btn-dl-folder" title="Download as ZIP">&#x2B07; Download</button>` : ""}
        ${_canRenameItems() ? `<button class="btn-action btn-rename" title="Rename">✏️ Rename</button>` : ""}
        ${_canDeleteItems() ? `<button class="btn-action btn-action-danger btn-delete" title="Delete">🗑️ Delete</button>` : ""}
      </div>
    </td>`;
  tr.querySelector(".file-link").addEventListener("click", (e) => {
    e.preventDefault();
    _loadFiles(folder.name);
  });
  tr.querySelector(".row-chk").addEventListener("change", (e) => {
    if (e.target.checked) _selection.add(folder.name); else _selection.delete(folder.name);
    tr.classList.toggle("row-selected", e.target.checked);
    _updateSelectionBar();
  });
  tr.querySelector(".btn-props").addEventListener("click", () => _showFolderProperties(folder));
  tr.querySelector(".btn-copy-url").addEventListener("click", (e) => _showCopyMenu(e.currentTarget, folder, "folder"));
  if (_canSas()) tr.querySelector(".btn-sas").addEventListener("click", () => _showSasModal(folder, true));
  if (CONFIG.app.allowDownload) tr.querySelector(".btn-dl-folder").addEventListener("click", () => _handleFolderDownload(folder));
  if (_canRenameItems()) tr.querySelector(".btn-rename").addEventListener("click", () => _showRenameModal(folder, "folder"));
  if (_canDeleteItems()) tr.querySelector(".btn-delete").addEventListener("click", () => _showDeleteModal(folder, "folder"));
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
        <span class="file-icon">${getFileIcon(file.displayName)}</span>
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
        <button class="btn-action btn-props"    title="Properties">ℹ️ Info</button>
        <button class="btn-action btn-copy-url" title="Copy URL">🔗 Copy URL</button>
        ${_canSas() ? `<button class="btn-action btn-sas" title="Generate SAS">🔑 SAS</button>` : ""}
        ${_isViewable(file.displayName) ? `<button class="btn-action btn-view" title="View file">👁 View</button>` : ""}
        ${_isViewable(file.displayName) && _canEditItems() ? `<button class="btn-action btn-edit" title="Edit file">📝 Edit</button>` : ""}
        ${CONFIG.app.allowDownload ? `<button class="btn-action btn-dl" title="Download">&#x2B07; Download</button>` : ""}
        ${_canRenameItems() ? `<button class="btn-action btn-rename" title="Rename">✏️ Rename</button>` : ""}
        ${_canDeleteItems() ? `<button class="btn-action btn-action-danger btn-delete" title="Delete">🗑️ Delete</button>` : ""}
      </div>
    </td>`;
  tr.querySelector(".btn-props").addEventListener("click", () => _showProperties(file));
  tr.querySelector(".row-chk").addEventListener("change", (e) => {
    if (e.target.checked) _selection.add(file.name); else _selection.delete(file.name);
    tr.classList.toggle("row-selected", e.target.checked);
    _updateSelectionBar();
  });
  tr.querySelector(".btn-copy-url").addEventListener("click", (e) => _showCopyMenu(e.currentTarget, file, "file"));
  if (_canSas()) tr.querySelector(".btn-sas").addEventListener("click", () => _showSasModal(file, false));
  if (_isViewable(file.displayName)) tr.querySelector(".btn-view").addEventListener("click", () => _showViewModal(file));
  if (_isViewable(file.displayName) && _canEditItems()) tr.querySelector(".btn-edit").addEventListener("click", () => _showEditModal(file));
  if (CONFIG.app.allowDownload) tr.querySelector(".btn-dl").addEventListener("click", () => _handleDownload(file));
  if (_canRenameItems()) tr.querySelector(".btn-rename").addEventListener("click", () => _showRenameModal(file, "file"));
  if (_canDeleteItems()) tr.querySelector(".btn-delete").addEventListener("click", () => _showDeleteModal(file, "file"));
  return tr;
}

// ── Search-results list view ──────────────────────────────────

/** List view for whole-container search results (flat rows, full blob paths). */
function _renderSearchResultsListView(container, files) {
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
        <span class="file-icon">${getFileIcon(fileName)}</span>
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
        ${parentPrefix ? `<button class="btn-action btn-goto-folder" title="Open containing folder">📂 Folder</button>` : ""}
        <button class="btn-action btn-props"    title="Properties">ℹ️ Info</button>
        <button class="btn-action btn-copy-url" title="Copy URL">🔗 Copy URL</button>
        ${_canSas() ? `<button class="btn-action btn-sas" title="Generate SAS">🔑 SAS</button>` : ""}
        ${_isViewable(fileName) ? `<button class="btn-action btn-view" title="View file">👁 View</button>` : ""}
        ${_isViewable(fileName) && _canEditItems() ? `<button class="btn-action btn-edit" title="Edit file">📝 Edit</button>` : ""}
        ${CONFIG.app.allowDownload ? `<button class="btn-action btn-dl" title="Download">&#x2B07; Download</button>` : ""}
        ${_canRenameItems() ? `<button class="btn-action btn-rename" title="Rename">✏️ Rename</button>` : ""}
        ${_canDeleteItems() ? `<button class="btn-action btn-action-danger btn-delete" title="Delete">🗑️ Delete</button>` : ""}
      </div>
    </td>`;

  tr.querySelector(".row-chk").addEventListener("change", (e) => {
    if (e.target.checked) _selection.add(file.name); else _selection.delete(file.name);
    tr.classList.toggle("row-selected", e.target.checked);
    _updateSelectionBar();
  });
  if (parentPrefix) {
    tr.querySelector(".btn-goto-folder").addEventListener("click", () => {
      _clearSearch();
      _loadFiles(parentPrefix);
    });
  }
  tr.querySelector(".btn-props").addEventListener("click", () => _showProperties(file));
  tr.querySelector(".btn-copy-url").addEventListener("click", (e) => _showCopyMenu(e.currentTarget, file, "file"));
  if (_canSas()) tr.querySelector(".btn-sas").addEventListener("click", () => _showSasModal(file, false));
  if (_isViewable(fileName)) tr.querySelector(".btn-view").addEventListener("click", () => _showViewModal(file));
  if (_isViewable(fileName) && _canEditItems()) tr.querySelector(".btn-edit").addEventListener("click", () => _showEditModal(file));
  if (CONFIG.app.allowDownload) tr.querySelector(".btn-dl").addEventListener("click", () => _handleDownload(file));
  if (_canRenameItems()) tr.querySelector(".btn-rename").addEventListener("click", () => _showRenameModal(file, "file"));
  if (_canDeleteItems()) tr.querySelector(".btn-delete").addEventListener("click", () => _showDeleteModal(file, "file"));
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
      <button class="btn-action btn-props"    title="Properties">ℹ️</button>
      <button class="btn-action btn-copy-url" title="Copy URL">🔗</button>
      ${_canSas() ? `<button class="btn-action btn-sas" title="Generate SAS">🔑</button>` : ""}
      ${CONFIG.app.allowDownload ? `<button class="btn-action btn-dl-folder" title="Download as ZIP">&#x2B07;</button>` : ""}
      ${_canRenameItems() ? `<button class="btn-action btn-rename" title="Rename">✏️</button>` : ""}
      ${_canDeleteItems() ? `<button class="btn-action btn-action-danger btn-delete" title="Delete">🗑️</button>` : ""}
    </div>`;
  div.addEventListener("click", (e) => {
    if (e.target.closest(".btn-action") || e.target.closest(".card-chk")) return;
    _loadFiles(folder.name);
  });
  div.addEventListener("keydown", (e) => {
    if (e.key === "Enter" || e.key === " ") { e.preventDefault(); _loadFiles(folder.name); }
  });
  div.querySelector(".card-chk").addEventListener("change", (e) => {
    e.stopPropagation();
    if (e.target.checked) _selection.add(folder.name); else _selection.delete(folder.name);
    div.classList.toggle("card-selected", e.target.checked);
    _updateSelectionBar();
  });
  div.querySelector(".btn-props").addEventListener("click", (e) => { e.stopPropagation(); _showFolderProperties(folder); });
  div.querySelector(".btn-copy-url").addEventListener("click", (e) => { e.stopPropagation(); _showCopyMenu(e.currentTarget, folder, "folder"); });
  if (_canSas()) div.querySelector(".btn-sas").addEventListener("click", (e) => { e.stopPropagation(); _showSasModal(folder, true); });
  if (CONFIG.app.allowDownload) div.querySelector(".btn-dl-folder").addEventListener("click", (e) => { e.stopPropagation(); _handleFolderDownload(folder); });
  if (_canRenameItems()) div.querySelector(".btn-rename").addEventListener("click", (e) => { e.stopPropagation(); _showRenameModal(folder, "folder"); });
  if (_canDeleteItems()) div.querySelector(".btn-delete").addEventListener("click", (e) => { e.stopPropagation(); _showDeleteModal(folder, "folder"); });
  return div;
}

function _makeFileCard(file) {
  const div = document.createElement("div");
  div.className = "file-card";
  if (_selection.has(file.name)) div.classList.add("card-selected");
  div.innerHTML = `
    <input type="checkbox" class="card-chk" data-name="${_esc(file.name)}" ${_selection.has(file.name) ? "checked" : ""} title="Select" />
    <div class="card-icon">${getFileIcon(file.displayName)}</div>
    <div class="card-name" title="${_esc(file.name)}">${_esc(file.displayName)}</div>
    <div class="card-meta">${formatFileSize(file.size)}</div>
    <div class="card-meta card-date">${file.createdOn ? `Created ${formatDateShort(file.createdOn)}` : ""}</div>
    <div class="card-actions">
      <button class="btn-action btn-props"    title="Properties">ℹ️</button>
      <button class="btn-action btn-copy-url" title="Copy URL">🔗</button>
      ${_canSas() ? `<button class="btn-action btn-sas" title="Generate SAS">🔑</button>` : ""}
      ${_isViewable(file.displayName) ? `<button class="btn-action btn-view" title="View file">👁</button>` : ""}
      ${_isViewable(file.displayName) && _canEditItems() ? `<button class="btn-action btn-edit" title="Edit file">📝</button>` : ""}
      ${CONFIG.app.allowDownload ? `<button class="btn-action btn-dl" title="Download">&#x2B07;</button>` : ""}
      ${_canRenameItems() ? `<button class="btn-action btn-rename" title="Rename">✏️</button>` : ""}
      ${_canDeleteItems() ? `<button class="btn-action btn-action-danger btn-delete" title="Delete">🗑️</button>` : ""}
    </div>`;
  div.querySelector(".card-chk").addEventListener("change", (e) => {
    e.stopPropagation();
    if (e.target.checked) _selection.add(file.name); else _selection.delete(file.name);
    div.classList.toggle("card-selected", e.target.checked);
    _updateSelectionBar();
  });
  div.querySelector(".btn-props").addEventListener("click", (e) => { e.stopPropagation(); _showProperties(file); });
  div.querySelector(".btn-copy-url").addEventListener("click", (e) => { e.stopPropagation(); _showCopyMenu(e.currentTarget, file, "file"); });
  if (_canSas()) div.querySelector(".btn-sas").addEventListener("click", (e) => { e.stopPropagation(); _showSasModal(file, false); });
  if (_isViewable(file.displayName)) div.querySelector(".btn-view").addEventListener("click", (e) => { e.stopPropagation(); _showViewModal(file); });
  if (_isViewable(file.displayName) && _canEditItems()) div.querySelector(".btn-edit").addEventListener("click", (e) => { e.stopPropagation(); _showEditModal(file); });
  if (CONFIG.app.allowDownload) div.querySelector(".btn-dl").addEventListener("click", (e) => { e.stopPropagation(); _handleDownload(file); });
  if (_canRenameItems()) div.querySelector(".btn-rename").addEventListener("click", (e) => { e.stopPropagation(); _showRenameModal(file, "file"); });
  if (_canDeleteItems()) div.querySelector(".btn-delete").addEventListener("click", (e) => { e.stopPropagation(); _showDeleteModal(file, "file"); });
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
  const lower = filename.toLowerCase();
  // Match files with no extension that are known config files
  if (["dockerfile","makefile",".env",".gitignore",".gitattributes"].includes(lower)) return true;
  const ext = lower.split(".").pop();
  return _VIEWABLE_EXTENSIONS.has(ext);
}

const _MAX_VIEW_BYTES = 2 * 1024 * 1024; // 2 MB preview cap

async function _showViewModal(file) {
  const modal = _el("viewModal");
  const body  = _el("viewModalBody");

  // Reset state
  body.innerHTML = `<div class="view-loading"><span class="view-spinner"></span> Loading…</div>`;
  _el("viewModalTitle").textContent = file.displayName;
  const ext = file.displayName.split(".").pop().toLowerCase();
  _el("viewLangBadge").textContent = ext.toUpperCase();
  _el("viewFileMeta").textContent  = "";
  _el("viewWrapBtn").classList.remove("active");
  modal.classList.remove("hidden");

  let content = null;
  let wrapped  = false;

  const close = () => { modal.classList.add("hidden"); body.innerHTML = ""; };
  _el("viewModalClose").onclick = close;
  _el("viewCloseBtn").onclick   = close;
  modal.onclick = (e) => { if (e.target === modal) close(); };

  try {
    const { accountName, containerName } = CONFIG.storage;
    const token = await getStorageToken();
    const url   = `https://${accountName}.blob.core.windows.net/${containerName}/${_encodePath(file.name)}`;

    const res = await fetch(url, {
      headers: { Authorization: `Bearer ${token}`, "x-ms-version": "2020-10-02" },
    });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);

    // Check size before reading body
    const cl = parseInt(res.headers.get("content-length") || "0", 10);
    if (cl > _MAX_VIEW_BYTES) {
      body.innerHTML = `<div class="view-too-large">
        <div style="font-size:32px">&#x26A0;&#xFE0F;</div>
        <p>File is too large to preview (${formatFileSize(cl)}).</p>
        ${CONFIG.app.allowDownload ? `<button id="viewDlInstead" class="btn btn-primary-sm">⬇ Download instead</button>` : ""}
      </div>`;
      if (CONFIG.app.allowDownload) _el("viewDlInstead").onclick = () => { close(); _handleDownload(file); };
      return;
    }

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
    _el("viewDlBtn").classList.toggle("hidden", !CONFIG.app.allowDownload);
    _el("viewDlBtn").onclick = () => _handleDownload(file);

  } catch (err) {
    body.innerHTML = `<div class="view-too-large" style="color:var(--error)">Failed to load file: ${_esc(err.message)}</div>`;
  }
}

function _renderViewContent(body, content, wrapped) {
  // Escape HTML in content for safe injection into innerHTML
  const safe = content
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");

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
    const token = await getStorageToken();
    const url   = `https://${accountName}.blob.core.windows.net/${containerName}/${_encodePath(file.name)}`;
    const res   = await fetch(url, {
      headers: { Authorization: `Bearer ${token}`, "x-ms-version": "2020-10-02" },
    });
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
    metaEl.textContent = `Failed to load: ${_esc(err.message)}`;
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
      try { meta = await getBlobMetadata(file.name); } catch { /* ignore */ }
      const upn = user?.username || "";
      const oid = user?.oid      || "";
      if (upn) meta.last_edited_by_upn = upn;
      if (oid) meta.last_edited_by_oid = oid;

      await uploadBlob(file.name, newFile, null, meta);
      close();
      _loadFiles(_currentPrefix);
      _showToast(`✅ "${_esc(file.displayName)}" saved`);
    } catch (err) {
      metaEl.textContent = `Save failed: ${_esc(err.message)}`;
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

// ── Download ─────────────────────────────────────────────────

// ── CSV Report ────────────────────────────────────────────

async function _exportReport() {
  _showToast("⏳ Scanning… building report", 60000);
  _showLoading(true);
  try {
    const { accountName, containerName } = CONFIG.storage;
    const rows = [];

    // CSV header
    rows.push([
      "Type",
      "Name",
      "Full Path",
      "Path Length",
      "Blob URL",
      "Size (MB)",
      "Size (bytes)",
      "Content Type",
      "Last Modified",
      "Created On",
      "ETag",
      "MD5",
    ]);

    // Recursively collect every folder and file under the current prefix
    async function collect(prefix) {
      const { folders, files } = await listBlobsAtPrefix(prefix);

      for (const folder of folders) {
        const encoded = folder.name.split("/").map(encodeURIComponent).join("/");
        const blobUrl = `https://${accountName}.blob.core.windows.net/${containerName}/${encoded}`;
        rows.push([
          "folder",
          folder.displayName,
          folder.name,
          String(folder.name.length),
          blobUrl,
          "",   // no size for folders
          "",
          "",
          "",
          "",
          "",
          "",
        ]);
        await collect(folder.name);
      }

      for (const file of files) {
        const encoded = file.name.split("/").map(encodeURIComponent).join("/");
        const blobUrl = `https://${accountName}.blob.core.windows.net/${containerName}/${encoded}`;
        const sizeMB  = file.size > 0
          ? (file.size / (1024 * 1024)).toFixed(6)
          : "0";
        rows.push([
          "file",
          file.displayName,
          file.name,
          String(file.name.length),
          blobUrl,
          sizeMB,
          String(file.size),
          file.contentType || "",
          file.lastModified || "",
          file.createdOn   || "",
          file.etag        || "",
          file.md5         || "",
        ]);
      }
    }

    await collect(_currentPrefix);

    // Serialize to CSV (RFC 4180 — fields with commas/quotes/newlines are quoted)
    const csvContent = rows.map(row =>
      row.map(cell => {
        const s = String(cell ?? "");
        return (s.includes(",") || s.includes('"') || s.includes("\n"))
          ? `"${s.replace(/"/g, '""')}"`
          : s;
      }).join(",")
    ).join("\r\n");

    // Trigger download
    const now      = new Date().toISOString().slice(0, 19).replace(/[T:]/g, "-");
    const location = _currentPrefix
      ? _currentPrefix.replace(/\/$/, "").split("/").pop()
      : containerName;
    const filename = `report_${location}_${now}.csv`;

    const blob   = new Blob(["\uFEFF" + csvContent], { type: "text/csv;charset=utf-8;" });
    const url    = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href     = url;
    anchor.download = filename;
    document.body.appendChild(anchor);
    anchor.click();
    document.body.removeChild(anchor);
    URL.revokeObjectURL(url);

    _showToast(`✅ ${filename} — ${rows.length - 1} items exported`);
  } catch (err) {
    console.error("[report]", err);
    _showError(`Report failed: ${err.message}`);
  } finally {
    _showLoading(false);
  }
}

// ── Download ────────────────────────────────────────────

async function _handleDownload(file) {
  _showLoading(true);
  try {
    await downloadBlob(file.name);
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

function _showCopyMenu(btn, item, kind) {
  // Close any already-open menu
  if (_copyMenuOpen) { _copyMenuOpen.remove(); _copyMenuOpen = null; }

  const { accountName, containerName } = CONFIG.storage;
  const encoded = item.name.split("/").map(encodeURIComponent).join("/");
  const blobUrl = `https://${accountName}.blob.core.windows.net/${containerName}/${encoded}`;
  const appUrl  = `${window.location.origin}${window.location.pathname}#${item.name}`;

  const menu = document.createElement("div");
  menu.className = "copy-menu";

  [
    { icon: "🗄️", label: "Blob URL",  sub: "Direct link to the file in Azure Storage",        url: blobUrl, isAppLink: false },
    { icon: "🔗",  label: "App link", sub: "Link that opens this explorer at this location", url: appUrl,  isAppLink: true  },
  ].forEach(({ icon, label, sub, url, isAppLink }) => {
    const menuItem = document.createElement("button");
    menuItem.type = "button";
    menuItem.className = "copy-menu-item";
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
    menuItem.appendChild(iconSpan);
    menuItem.appendChild(textSpan);
    menuItem.addEventListener("click", (e) => {
      e.stopPropagation();
      menu.remove();
      _copyMenuOpen = null;
      if (isAppLink && CONFIG.app.allowDownload) {
        // Copy the URL, then offer a one-click download action in the toast
        const dlFn = kind === "folder"
          ? () => _handleFolderDownload(item)
          : () => _handleDownload(item);
        const doToast = () => _showToast("📋 App link copied!", 8000, "⬇ Download", dlFn);
        if (navigator.clipboard && window.isSecureContext) {
          navigator.clipboard.writeText(url).then(doToast).catch(() => { _fallbackCopy(url); doToast(); });
        } else {
          _fallbackCopy(url);
          doToast();
        }
      } else {
        _copyToClipboard(url);
      }
    });
    menu.appendChild(menuItem);
  });

  document.body.appendChild(menu);
  _copyMenuOpen = menu;

  // Use fixed positioning — no scroll offset math needed
  const rect = btn.getBoundingClientRect();
  const menuW = menu.offsetWidth || 260;
  let left = rect.left;
  if (left + menuW > window.innerWidth - 8) left = window.innerWidth - menuW - 8;
  menu.style.top  = `${rect.bottom + 4}px`;
  menu.style.left = `${Math.max(8, left)}px`;

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
      errEl.textContent = "Name cannot contain \"\\\"";
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
        // Rename all blobs sharing the folder prefix
        const { folders, files } = await listBlobsAtPrefix(srcName);
        const allBlobs = [
          ...files.map(f => f.name),
          ...folders.map(f => f.name),
        ];
        // If folder is empty just reflect the new name (virtual folders don\'t exist as blobs)
        if (allBlobs.length === 0) {
          // No real blobs to move — just close and refresh
        } else {
          for (const blob of allBlobs) {
            const rel  = blob.slice(srcName.length);
            const dest = destName + rel;
            await renameBlob(blob, dest);
          }
        }
      } else {
        await renameBlob(srcName, destName);
      }
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
  _el("selBulkDeleteBtn").addEventListener("click",   _bulkDelete);
});

async function _bulkDownload() {
  if (_selection.size === 0) return;
  _showLoading(true);
  _showToast(`⏳ Building ZIP for ${_selection.size} item${_selection.size !== 1 ? "s" : ""}…`);
  try {
    const { accountName, containerName } = CONFIG.storage;
    const token = await getStorageToken();

    // Recursively collect all blob names for a name (file or folder)
    async function collectNames(name) {
      if (name.endsWith("/")) {
        // It's a folder — recurse
        const { folders, files } = await listBlobsAtPrefix(name);
        let names = files.map(f => f.name);
        for (const sub of folders) names = names.concat(await collectNames(sub.name));
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

    const entries = [];
    for (const name of allNames) {
      const url = `https://${accountName}.blob.core.windows.net/${containerName}/${_encodePath(name)}`;
      const res = await fetch(url, {
        headers: { Authorization: `Bearer ${token}`, "x-ms-version": "2020-10-02" },
      });
      if (!res.ok) throw new Error(`Failed to fetch "${name}" (${res.status})`);
      const data = new Uint8Array(await res.arrayBuffer());
      entries.push({ name: name.startsWith(stripPrefix) ? name.slice(stripPrefix.length) : name, data });
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
      _selection.clear();
      _updateSelectionBar();
      close();
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
      close();
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
        close();
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

  const addRow = (label, value, muted = false) => {
    const tr = document.createElement("tr");
    const tdLabel = document.createElement("td"); tdLabel.textContent = label;
    const tdVal   = document.createElement("td"); tdVal.innerHTML = value;
    if (muted) tdVal.style.color = "var(--text-muted)";
    tr.appendChild(tdLabel); tr.appendChild(tdVal);
    table.appendChild(tr);
  };

  addRow("Account",   _esc(CONFIG.storage.accountName));
  addRow("Container", _esc(CONFIG.storage.containerName));
  addRow("Path",      `<code style="font-size:12px">${_esc(displayPath)}</code>`);

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
    addRow("Error", _esc(e.message));
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
      if (!overwrite && await blobExists(item.blobPath)) {
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
  if (pending.length > 0) _loadFiles(_currentPrefix);
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
  try { meta = await getBlobMetadata(blobPath); } catch { /* ignore */ }
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
    <span class="upload-item-icon">${getFileIcon(item.file.name)}</span>
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

  return node;
}

async function _loadTreeChildren(prefix, parentNode) {
  const childrenEl = parentNode.querySelector(":scope > .tree-children");
  childrenEl.innerHTML = `<div class="tree-loading">Loading\u2026</div>`;

  try {
    const { folders } = await listBlobsAtPrefix(prefix);
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
  } catch {
    childrenEl.innerHTML = `<div class="tree-loading tree-load-err">Failed to load</div>`;
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

function _hideNetworkErrorHelp() {
  _el("networkErrorHelp").classList.add("hidden");
}

function _el(id) {
  return document.getElementById(id);
}

function _esc(text) {
  const d = document.createElement("div");
  d.textContent = text;
  return d.innerHTML;
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
  const closeBtn    = _el("sasModalClose");

  const close = () => modal.classList.add("hidden");
  closeBtn.onclick  = close;
  cancelBtn.onclick = close;
  modal.onclick = (e) => { if (e.target === modal) close(); };

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
      resultEl.select();
      document.execCommand("copy");
    }
  };
}