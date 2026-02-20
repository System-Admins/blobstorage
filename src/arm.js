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

/**
 * List all storage accounts in a subscription.
 * @param {string} subscriptionId
 * @returns {Promise<Array<{id:string, name:string, location:string, resourceGroup:string, kind:string}>>}
 */
async function listStorageAccounts(subscriptionId) {
  const token = await getArmToken();
  const res   = await fetch(
    `${_ARM_BASE}/subscriptions/${subscriptionId}/providers/Microsoft.Storage/storageAccounts?api-version=${_ARM_VERSION}`,
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
  const url   = `${_ARM_BASE}/subscriptions/${subscriptionId}/resourceGroups/${resourceGroup}/providers/Microsoft.Storage/storageAccounts/${accountName}/blobServices/default/containers?api-version=${_ARM_VERSION}`;
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
 * Piggybacks on auth.js's refresh token.
 * @returns {Promise<string>}
 */
async function getArmToken() {
  // Return cached token if still valid (>60s remaining)
  const cached = sessionStorage.getItem(_ARM_TOKEN_KEY);
  const expiry  = parseInt(sessionStorage.getItem(_ARM_EXPIRY_KEY) || "0", 10);
  if (cached && Date.now() < expiry) return cached;

  // Fetch a new token using the stored refresh token
  const rt = sessionStorage.getItem(_KEYS.REFRESH_TOKEN);
  if (!rt) throw new Error("No refresh token — please sign in again.");

  const res = await fetch(
    `https://login.microsoftonline.com/${CONFIG.auth.tenantId}/oauth2/v2.0/token`,
    {
      method:  "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body:    new URLSearchParams({
        client_id:     CONFIG.auth.clientId,
        grant_type:    "refresh_token",
        refresh_token: rt,
        scope:         _ARM_SCOPE + " offline_access",
      }).toString(),
    }
  );

  const data = await res.json();
  if (!res.ok) throw new Error(data.error_description || data.error || "ARM token request failed");

  sessionStorage.setItem(_ARM_TOKEN_KEY,  data.access_token);
  sessionStorage.setItem(_ARM_EXPIRY_KEY, String(Date.now() + (data.expires_in - 60) * 1000));
  // Rotate refresh token if Microsoft returned a new one
  if (data.refresh_token) sessionStorage.setItem(_KEYS.REFRESH_TOKEN, data.refresh_token);

  return data.access_token;
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
    const url   = `${_ARM_BASE}/subscriptions/${subscriptionId}/resourceGroups/${resourceGroup}`
                + `/providers/Microsoft.Storage/storageAccounts/${accountName}`
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
