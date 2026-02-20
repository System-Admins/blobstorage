// ============================================================
//  ARM — Azure Resource Manager API helpers
//  Discovers storage accounts and containers accessible to the
//  signed-in user via the ARM REST API.
//  Depends on: config.js, auth.js
// ============================================================

const _ARM_BASE    = "https://management.azure.com";
const _ARM_VERSION = "2023-01-01";
const _ARM_SCOPE   = "https://management.azure.com/user_impersonation";

// sessionStorage key for the ARM access token
const _ARM_TOKEN_KEY    = "be_arm_token";
const _ARM_EXPIRY_KEY   = "be_arm_token_expiry";

// ── Public API ───────────────────────────────────────────────

/**
 * List all Azure subscriptions the signed-in user has access to.
 * @returns {Promise<Array<{id:string, displayName:string, subscriptionId:string}>>}
 */
async function listSubscriptions() {
  const token = await getArmToken();
  const res   = await fetch(
    `${_ARM_BASE}/subscriptions?api-version=2022-12-01`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  if (!res.ok) throw new Error(await _armError(res));
  const data = await res.json();
  return (data.value || []).filter(s => s.state === "Enabled");
}

/** Encode a single ARM path segment to prevent path-traversal or injection. */
function _encArmSegment(segment) {
  return encodeURIComponent(segment);
}

/**
 * List all storage accounts in a subscription.
 * @param {string} subscriptionId
 * @returns {Promise<Array<{id:string, name:string, location:string, resourceGroup:string, kind:string}>>}
 */
async function listStorageAccounts(subscriptionId) {
  const token = await getArmToken();
  const res   = await fetch(
    `${_ARM_BASE}/subscriptions/${_encArmSegment(subscriptionId)}/providers/Microsoft.Storage/storageAccounts?api-version=${_ARM_VERSION}`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  if (!res.ok) throw new Error(await _armError(res));
  const data = await res.json();
  return (data.value || []).map(a => ({
    id:            a.id,
    name:          a.name,
    location:      a.location,
    resourceGroup: _parseRg(a.id),
    kind:          a.kind,
    subscriptionId,
  }));
}

/**
 * List all blob containers in a storage account.
 * @param {string} subscriptionId
 * @param {string} resourceGroup
 * @param {string} accountName
 * @returns {Promise<Array<{name:string, publicAccess:string, leaseState:string}>>}
 */
async function listContainers(subscriptionId, resourceGroup, accountName) {
  const token = await getArmToken();
  const url   = `${_ARM_BASE}/subscriptions/${_encArmSegment(subscriptionId)}/resourceGroups/${_encArmSegment(resourceGroup)}/providers/Microsoft.Storage/storageAccounts/${_encArmSegment(accountName)}/blobServices/default/containers?api-version=${_ARM_VERSION}`;
  const res   = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  if (!res.ok) throw new Error(await _armError(res));
  const data = await res.json();
  return (data.value || []).map(c => ({
    name:         c.name,
    publicAccess: c.properties?.publicAccess || "None",
    leaseState:   c.properties?.leaseState   || "",
  }));
}

// ── ARM token ────────────────────────────────────────────────

/**
 * Return a valid ARM access token, refreshing silently if expired.
 * Piggybacks on auth.js's refresh token via the shared _refreshTokenForScope() helper.
 * @returns {Promise<string>}
 */
async function getArmToken() {
  return _refreshTokenForScope(
    _ARM_SCOPE + " offline_access",
    _ARM_TOKEN_KEY,
    _ARM_EXPIRY_KEY,
    "ARM"
  );
}

// ── Helpers ──────────────────────────────────────────────────

/**
 * Check whether the blob service of a storage account has a CORS rule that
 * allows the current browser origin.
 * Returns true  if a matching rule is found (or if the ARM call fails —
 *               we don't want to hide accounts just because the ARM CORS
 *               endpoint returned 403 due to insufficient IAM permissions).
 * Returns false only when the account is definitively reachable but has
 * zero CORS rules that permit the current origin.
 *
 * @param {string} subscriptionId
 * @param {string} resourceGroup
 * @param {string} accountName
 * @returns {Promise<boolean>}
 */
async function getBlobServiceCors(subscriptionId, resourceGroup, accountName) {
  try {
    const token = await getArmToken();
    const url   = `${_ARM_BASE}/subscriptions/${_encArmSegment(subscriptionId)}/resourceGroups/${_encArmSegment(resourceGroup)}`
                + `/providers/Microsoft.Storage/storageAccounts/${_encArmSegment(accountName)}`
                + `/blobServices/default?api-version=${_ARM_VERSION}`;
    const res   = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
    if (!res.ok) return true; // Can't confirm — keep account visible
    const data  = await res.json();
    const rules = data?.properties?.cors?.corsRules ?? [];
    if (rules.length === 0) return false; // No rules at all
    const origin = window.location.origin.toLowerCase();
    return rules.some(rule => {
      const origins = (rule.allowedOrigins ?? []).map(o => o.toLowerCase());
      if (!origins.includes("*") && !origins.includes(origin)) return false;
      const methods = (rule.allowedMethods ?? []).map(m => m.toUpperCase());
      // Must allow at least GET (read) — wildcard (*) is not a standard CORS method value
      return methods.includes("GET");
    });
  } catch {
    return true; // Network/token error — keep account visible
  }
}

function _parseRg(resourceId) {
  // /subscriptions/{sub}/resourceGroups/{rg}/providers/...
  const m = resourceId.match(/resourceGroups\/([^/]+)\//i);
  return m ? m[1] : "";
}

async function _armError(res) {
  try {
    const j = await res.json();
    return j?.error?.message || `ARM API ${res.status}`;
  } catch {
    return `ARM API ${res.status}`;
  }
}
