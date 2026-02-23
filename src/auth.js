// ============================================================
//  AUTH — OAuth 2.0 Authorization Code Flow + PKCE
//  Pure browser implementation — no external libraries required
//  Uses: fetch, crypto.subtle, sessionStorage, window.history
//  Depends on: config.js
// ============================================================

// ── Storage keys ─────────────────────────────────────────────
// All tokens (including refresh token) are stored in sessionStorage so they
// are scoped to the browser tab and cleared when it closes — this limits
// the blast radius of any XSS. Users must re-authenticate per session.

const _KEYS = {
  ACCESS_TOKEN:  "be_access_token",
  REFRESH_TOKEN: "be_refresh_token",
  TOKEN_EXPIRY:  "be_token_expiry",
  USER:          "be_user",
  CODE_VERIFIER: "be_pkce_verifier",
  STATE:         "be_oauth_state",
};

// Scopes for token requests — one resource audience at a time
const _SCOPES = "https://storage.azure.com/user_impersonation offline_access openid profile";

// All scopes included in the *authorization* URL so the user consents to both
// Storage and ARM in a single sign-in prompt. Token exchanges still use the
// resource-specific _SCOPES — the ARM token is fetched separately via the
// refresh token once consent has been granted.
// Mail.Send is included so the user can share links via email (Exchange Online).
const _SCOPES_AUTHORIZE = _SCOPES + " https://management.azure.com/user_impersonation https://graph.microsoft.com/Mail.Send";

// ── Endpoint helpers ─────────────────────────────────────────

const _base = () =>
  `https://login.microsoftonline.com/${CONFIG.auth.tenantId}/oauth2/v2.0`;

// ── Public API ───────────────────────────────────────────────

/**
 * Call once on page load.
 * • If the URL contains ?code=..., exchanges it for tokens (post-redirect).
 * • If a valid access token is stored, returns the cached user.
 * • If only a refresh token is stored, refreshes silently.
 * • Otherwise returns null — caller should show the sign-in page.
 * @returns {Promise<{name:string, username:string}|null>}
 */
async function initAuth() {
  const params = new URLSearchParams(window.location.search);
  const code   = params.get("code");
  const state  = params.get("state");
  const error  = params.get("error");

  // ── Returning from Microsoft login redirect ──
  if (error) {
    _cleanUrl();
    // If this was a silent (prompt=none) attempt, a login_required / interaction_required
    // error just means no active SSO session — fall through to the sign-in page.
    const wasSilent = sessionStorage.getItem("be_silent_attempt");
    sessionStorage.removeItem("be_silent_attempt");
    if (wasSilent && /login_required|interaction_required|consent_required/i.test(error)) {
      return null; // Let the caller show the sign-in page
    }
    throw new Error(params.get("error_description") || error);
  }

  if (code) {
    const savedState   = sessionStorage.getItem(_KEYS.STATE);
    const codeVerifier = sessionStorage.getItem(_KEYS.CODE_VERIFIER);
    sessionStorage.removeItem(_KEYS.STATE);
    sessionStorage.removeItem(_KEYS.CODE_VERIFIER);
    _cleanUrl();

    if (state !== savedState) {
      throw new Error("OAuth state mismatch — possible CSRF attack. Please sign in again.");
    }

    if (!codeVerifier) {
      throw new Error("Missing PKCE code verifier — please sign in again.");
    }

    await _exchangeCode(code, codeVerifier);
    return getUser();
  }

  // ── Already have a valid access token ──
  if (_hasValidToken()) {
    return getUser();
  }

  // ── Try a silent refresh ──
  const rt = sessionStorage.getItem(_KEYS.REFRESH_TOKEN);
  if (rt) {
    try {
      await _doRefresh(rt);
      return getUser();
    } catch (err) {
      console.warn("[auth] Silent refresh failed:", err.message);
      _clearSession();
    }
  }

  return null; // Caller must show sign-in UI
}

/**
 * Attempt a silent SSO sign-in using prompt=none.
 * The browser redirects away. If an active Entra ID session exists the user
 * comes back already authenticated; if not, initAuth() will return null and
 * the sign-in page is shown.
 */
async function signInSilent() {
  const verifier  = _generateVerifier();
  const challenge = await _generateChallenge(verifier);
  const state     = _generateState();

  sessionStorage.setItem(_KEYS.CODE_VERIFIER, verifier);
  sessionStorage.setItem(_KEYS.STATE, state);
  sessionStorage.setItem("be_silent_attempt", "1");
  _saveDeepLink();

  window.location.href = _base() + "/authorize?" + new URLSearchParams({
    client_id:             CONFIG.auth.clientId,
    response_type:         "code",
    redirect_uri:          CONFIG.auth.redirectUri,
    scope:                 _SCOPES_AUTHORIZE,
    code_challenge:        challenge,
    code_challenge_method: "S256",
    state,
    prompt:                "none", // succeed silently or return login_required
  });
}

/**
 * Redirect the browser to the Microsoft login page (Authorization Code + PKCE).
 * The function never returns — the page navigates away.
 * On return, initAuth() will pick up the authorization code automatically.
 */
async function signIn() {
  const verifier  = _generateVerifier();
  const challenge = await _generateChallenge(verifier);
  const state     = _generateState();

  sessionStorage.setItem(_KEYS.CODE_VERIFIER, verifier);
  sessionStorage.setItem(_KEYS.STATE, state);
  _saveDeepLink();

  window.location.href = _base() + "/authorize?" + new URLSearchParams({
    client_id:             CONFIG.auth.clientId,
    response_type:         "code",
    redirect_uri:          CONFIG.auth.redirectUri,
    scope:                 _SCOPES_AUTHORIZE,
    code_challenge:        challenge,
    code_challenge_method: "S256",
    state,
    // No prompt parameter — lets Microsoft reuse the existing SSO session
    // silently. If no session exists, the Microsoft login page is shown.
  });
  // Execution stops here — browser navigates away
}

/**
 * Clear tokens locally and redirect to Microsoft's logout endpoint.
 * The page navigates away and returns to redirectUri when done.
 */
function signOut() {
  _clearSession();
  window.location.href =
    _base() + "/logout?" +
    new URLSearchParams({ post_logout_redirect_uri: CONFIG.auth.redirectUri });
}

/**
 * Return a valid access token for Azure Storage, refreshing silently if needed.
 * Throws if the session has fully expired and the user must sign in again.
 * @returns {Promise<string>} Bearer access token
 */
async function getStorageToken() {
  if (_hasValidToken()) {
    return sessionStorage.getItem(_KEYS.ACCESS_TOKEN);
  }

  const rt = sessionStorage.getItem(_KEYS.REFRESH_TOKEN);
  if (rt) {
    await _doRefresh(rt);
    return sessionStorage.getItem(_KEYS.ACCESS_TOKEN);
  }

  throw new Error("Session expired. Please sign in again.");
}

/** Return the stored user object or null. */
function getUser() {
  const raw = sessionStorage.getItem(_KEYS.USER);
  if (!raw) return null;
  try {
    const u = JSON.parse(raw);
    if (typeof u?.name !== "string") return null;
    return { name: u.name, username: String(u.username || ""), oid: String(u.oid || "") };
  } catch {
    return null;
  }
}


// ── Microsoft Graph token ────────────────────────────────────

const _GRAPH_TOKEN_KEY  = "be_graph_token";
const _GRAPH_EXPIRY_KEY = "be_graph_token_expiry";
const _GRAPH_SCOPE      = "https://graph.microsoft.com/Mail.Send offline_access";

/**
 * Return a valid Microsoft Graph access token, refreshing silently if expired.
 * Piggybacks on auth.js's refresh token.
 * @returns {Promise<string>}
 */
async function getGraphToken() {
  return _refreshTokenForScope(_GRAPH_SCOPE, _GRAPH_TOKEN_KEY, _GRAPH_EXPIRY_KEY, "Graph");
}

/**
 * Shared helper: return a cached token for the given scope/resource, or
 * silently refresh it using the stored refresh token.
 * Used by getGraphToken() (auth.js) and getArmToken() (arm.js).
 *
 * @param {string} scope      OAuth scope string for the token request
 * @param {string} tokenKey   sessionStorage key for the access token
 * @param {string} expiryKey  sessionStorage key for the expiry timestamp
 * @param {string} label      Human-readable label for error messages (e.g. "Graph", "ARM")
 * @returns {Promise<string>}  Bearer access token
 */
async function _refreshTokenForScope(scope, tokenKey, expiryKey, label) {
  const cached = sessionStorage.getItem(tokenKey);
  const expiry = parseInt(sessionStorage.getItem(expiryKey) || "0", 10);
  if (cached && Date.now() < expiry) return cached;

  const rt = sessionStorage.getItem(_KEYS.REFRESH_TOKEN);
  if (!rt) throw new Error("No refresh token — please sign in again.");

  const res = await fetch(_base() + "/token", {
    method:  "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body:    new URLSearchParams({
      client_id:     CONFIG.auth.clientId,
      grant_type:    "refresh_token",
      refresh_token: rt,
      scope,
    }).toString(),
  });

  const data = await res.json();
  if (!res.ok) throw new Error(data.error_description || data.error || `${label} token request failed`);

  sessionStorage.setItem(tokenKey,  data.access_token);
  sessionStorage.setItem(expiryKey, String(Date.now() + (data.expires_in - 60) * 1000));
  if (data.refresh_token) sessionStorage.setItem(_KEYS.REFRESH_TOKEN, data.refresh_token);

  return data.access_token;
}

// ── Token requests ────────────────────────────────────────────

async function _exchangeCode(code, codeVerifier) {
  await _fetchTokens(new URLSearchParams({
    client_id:     CONFIG.auth.clientId,
    grant_type:    "authorization_code",
    code,
    redirect_uri:  CONFIG.auth.redirectUri,
    code_verifier: codeVerifier,
    scope:         _SCOPES,
  }));
}

async function _doRefresh(refreshToken) {
  await _fetchTokens(new URLSearchParams({
    client_id:     CONFIG.auth.clientId,
    grant_type:    "refresh_token",
    refresh_token: refreshToken,
    scope:         _SCOPES,
  }));
}

async function _fetchTokens(body) {
  const res  = await fetch(_base() + "/token", {
    method:  "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body:    body.toString(),
  });

  const data = await res.json();
  if (!res.ok) {
    throw new Error(data.error_description || data.error || `Token request failed (${res.status})`);
  }

  // Store access token with a 60-second early-expiry buffer
  sessionStorage.setItem(_KEYS.ACCESS_TOKEN, data.access_token);
  sessionStorage.setItem(
    _KEYS.TOKEN_EXPIRY,
    String(Date.now() + (data.expires_in - 60) * 1000)
  );

  // Persist refresh token in sessionStorage (tab-scoped — limits XSS exposure).
  // (returned on first exchange and when rotated)
  if (data.refresh_token) {
    sessionStorage.setItem(_KEYS.REFRESH_TOKEN, data.refresh_token);
  }

  // Decode user info from the id_token for display purposes
  // (signature is already verified server-side by Microsoft)
  if (data.id_token) {
    const claims = _parseJwt(data.id_token);
    sessionStorage.setItem(_KEYS.USER, JSON.stringify({
      name:     claims.name     || claims.preferred_username || "User",
      username: claims.preferred_username || claims.upn || claims.email || "",
      oid:      claims.oid || claims.sub || "",
    }));
  }
}

// ── PKCE / crypto helpers ─────────────────────────────────────

function _generateVerifier() {
  const buf = new Uint8Array(32);
  crypto.getRandomValues(buf);
  return _base64url(buf);
}

async function _generateChallenge(verifier) {
  const data   = new TextEncoder().encode(verifier);
  const digest = await crypto.subtle.digest("SHA-256", data);
  return _base64url(new Uint8Array(digest));
}

function _generateState() {
  const buf = new Uint8Array(16);
  crypto.getRandomValues(buf);
  return _base64url(buf);
}

/** Base64url-encode a Uint8Array (RFC 4648, no padding). */
function _base64url(buf) {
  return btoa(String.fromCharCode(...buf))
    .replace(/\+/g, "-")
    .replace(/\//g, "_")
    .replace(/=/g, "");
}

// ── JWT payload decoder ───────────────────────────────────────
// Used for display only — signature is verified server-side.

function _parseJwt(token) {
  try {
    const parts = token.split(".");
    if (parts.length !== 3) return {};
    const b64 = parts[1].replace(/-/g, "+").replace(/_/g, "/");
    const padLen = (4 - (b64.length % 4)) % 4;
    const padded = b64 + "=".repeat(padLen);
    const claims = JSON.parse(atob(padded));
    if (typeof claims !== "object" || claims === null) return {};
    return claims;
  } catch {
    return {};
  }
}

// ── Session helpers ───────────────────────────────────────────

function _hasValidToken() {
  const token  = sessionStorage.getItem(_KEYS.ACCESS_TOKEN);
  const expiry = parseInt(sessionStorage.getItem(_KEYS.TOKEN_EXPIRY) || "0", 10);
  return !!token && Date.now() < expiry;
}


function _clearSession() {
  Object.values(_KEYS).forEach((k) => sessionStorage.removeItem(k));
  // Also clear cached resource-specific tokens (ARM, Graph)
  sessionStorage.removeItem(_ARM_TOKEN_KEY);
  sessionStorage.removeItem(_ARM_EXPIRY_KEY);
  sessionStorage.removeItem(_GRAPH_TOKEN_KEY);
  sessionStorage.removeItem(_GRAPH_EXPIRY_KEY);
}

/** Remove ?code=...&state=... from the browser URL without a page reload. */
function _cleanUrl() {
  // Restore any deep-link hash that was saved before the auth redirect.
  // The hash is lost during the OAuth round-trip because the redirect_uri
  // only receives ?code=...&state=... with no hash fragment.
  const savedHash = sessionStorage.getItem("be_deep_link");
  sessionStorage.removeItem("be_deep_link");
  window.history.replaceState({}, document.title,
    window.location.pathname + (savedHash || ""));
}

/** Save the current URL hash so it survives an OAuth redirect round-trip. */
function _saveDeepLink() {
  if (window.location.hash) {
    sessionStorage.setItem("be_deep_link", window.location.hash);
  }
}
