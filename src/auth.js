// ============================================================
//  AUTH — OAuth 2.0 Authorization Code Flow + PKCE
//  Pure browser implementation — no external libraries required
//  Uses: fetch, crypto.subtle, sessionStorage, localStorage, window.history
//  Depends on: config.js
// ============================================================

// ── Storage keys ─────────────────────────────────────────────
// Access token, expiry, user info, PKCE verifier and OAuth state are stored in
// sessionStorage (cleared when the tab closes — limits XSS exposure).
// The refresh token is stored in localStorage so the user stays signed in
// across browser sessions until the token expires (~90 days for Entra ID).

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
const _SCOPES_AUTHORIZE = _SCOPES + " https://management.azure.com/user_impersonation";

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
    return _getUser();
  }

  // ── Already have a valid access token ──
  if (_hasValidToken()) {
    return _getUser();
  }

  // ── Try a silent refresh ──
  const rt = localStorage.getItem(_KEYS.REFRESH_TOKEN);
  if (rt) {
    try {
      await _doRefresh(rt);
      return _getUser();
    } catch (err) {
      console.warn("[auth] Silent refresh failed:", err.message);
      _clearSession();
    }
  }

  return null; // Caller must show sign-in UI
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

  const rt = localStorage.getItem(_KEYS.REFRESH_TOKEN);
  if (rt) {
    await _doRefresh(rt);
    return sessionStorage.getItem(_KEYS.ACCESS_TOKEN);
  }

  throw new Error("Session expired. Please sign in again.");
}

/** Return the stored user object or null. */
function getUser() {
  const raw = sessionStorage.getItem(_KEYS.USER);
  return raw ? JSON.parse(raw) : null;
}

/** Return true if a valid (non-expired) access token is currently stored. */
function isAuthenticated() {
  return _hasValidToken();
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

  // Persist refresh token in localStorage so the session survives tab/browser close.
  // (returned on first exchange and when rotated)
  if (data.refresh_token) {
    localStorage.setItem(_KEYS.REFRESH_TOKEN, data.refresh_token);
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
    const b64 = token.split(".")[1].replace(/-/g, "+").replace(/_/g, "/");
    const padLen = (4 - (b64.length % 4)) % 4;
    const padded = b64 + "=".repeat(padLen);
    return JSON.parse(atob(padded));
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

function _getUser() {
  return getUser();
}

function _clearSession() {
  Object.values(_KEYS).forEach((k) => sessionStorage.removeItem(k));
  localStorage.removeItem(_KEYS.REFRESH_TOKEN);
}

/** Remove ?code=...&state=... from the browser URL without a page reload. */
function _cleanUrl() {
  window.history.replaceState({}, document.title, window.location.pathname);
}
