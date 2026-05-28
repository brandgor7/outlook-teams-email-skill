/**
 * token.mjs -- Shared MSAL token helper
 *
 * Supports two account types:
 *   'personal' -- personal Microsoft account (consumers authority)
 *   'work'     -- work/school account (single-tenant via tenantId)
 *   'email'    -- prefers personal if configured, falls back to work
 *
 * Token files:
 *   personal: personal-tokens.json / personal-msal-cache.json
 *   work:     work-tokens.json     / work-msal-cache.json
 *
 * Legacy flat auth config (auth.clientId) is treated as 'personal' and falls
 * back to outlook-tokens.json / outlook-msal-cache.json for compatibility.
 */

import { readFileSync, writeFileSync, existsSync } from 'fs';
import { resolve, dirname } from 'path';
import { fileURLToPath } from 'url';
import { PublicClientApplication } from '@azure/msal-node';

const __dirname = dirname(fileURLToPath(import.meta.url));
const SKILL_ROOT = resolve(__dirname, '..');
const WORKSPACE_ROOT = resolve(SKILL_ROOT, '..', '..');
const CONFIG_PATH = resolve(SKILL_ROOT, 'config.json');

function resolveWorkspacePath(filename) {
  const wsPath = resolve(WORKSPACE_ROOT, filename);
  if (existsSync(wsPath)) return wsPath;
  return resolve(SKILL_ROOT, filename);
}

function loadConfig() {
  if (!existsSync(CONFIG_PATH)) {
    throw new Error(`config.json not found at ${CONFIG_PATH}. Copy config.json and fill in your credentials.`);
  }
  return JSON.parse(readFileSync(CONFIG_PATH, 'utf8'));
}

// Returns auth config {clientId, authority} for the given resolved account type.
function resolveAuthConfig(auth, account) {
  // Legacy flat format: auth.clientId at top level -> treat as personal
  if (auth?.clientId && account !== 'work') {
    return {
      clientId: auth.clientId,
      authority: auth.authority ?? 'https://login.microsoftonline.com/consumers',
    };
  }

  if (account === 'personal') {
    if (!auth?.personal?.clientId) {
      throw new Error(
        'auth.personal.clientId is not set in config.json.\n' +
        'Configure auth.personal and run: node scripts/auth.mjs --account=personal'
      );
    }
    return {
      clientId: auth.personal.clientId,
      authority: auth.personal.authority ?? 'https://login.microsoftonline.com/consumers',
    };
  }

  if (account === 'work') {
    if (!auth?.work?.clientId) {
      throw new Error(
        'auth.work.clientId is not set in config.json.\n' +
        'Configure auth.work and run: node scripts/auth.mjs --account=work'
      );
    }
    const tenantId = auth.work.tenantId;
    if (!tenantId && !auth.work.authority) {
      throw new Error('auth.work.tenantId is required in config.json for work account authentication.');
    }
    return {
      clientId: auth.work.clientId,
      authority: auth.work.authority ?? `https://login.microsoftonline.com/${tenantId}`,
    };
  }

  throw new Error(`Unknown account type: "${account}". Use 'personal' or 'work'.`);
}

// Returns token file paths for a resolved account type ('personal' | 'work').
function resolveTokenPaths(account) {
  if (account === 'work') {
    return {
      tokensPath: resolveWorkspacePath('work-tokens.json'),
      msalCachePath: resolveWorkspacePath('work-msal-cache.json'),
    };
  }
  // Personal: prefer new name, fall back to legacy outlook-tokens.json
  const primaryTokens = resolveWorkspacePath('personal-tokens.json');
  const primaryCache = resolveWorkspacePath('personal-msal-cache.json');
  const legacyTokens = resolveWorkspacePath('outlook-tokens.json');
  const legacyCache = resolveWorkspacePath('outlook-msal-cache.json');
  return {
    tokensPath: existsSync(primaryTokens) ? primaryTokens : (existsSync(legacyTokens) ? legacyTokens : primaryTokens),
    msalCachePath: existsSync(primaryCache) ? primaryCache : (existsSync(legacyCache) ? legacyCache : primaryCache),
  };
}

// Resolves 'email' to 'personal' or 'work' based on what is configured.
function resolveAccountType(auth, account) {
  if (account !== 'email') return account;
  if (auth?.clientId) return 'personal'; // legacy flat config
  if (auth?.personal?.clientId) return 'personal';
  if (auth?.work?.clientId) return 'work';
  throw new Error(
    'No auth config found in config.json.\n' +
    'Configure auth.personal and/or auth.work, then run:\n' +
    '  node scripts/auth.mjs --account=personal\n' +
    '  node scripts/auth.mjs --account=work'
  );
}

function buildMsalApp(authConfig, serializedCache) {
  const cachePlugin = {
    beforeCacheAccess: async (cacheContext) => {
      if (serializedCache) cacheContext.tokenCache.deserialize(serializedCache);
    },
    afterCacheAccess: async (_cacheContext) => {},
  };
  return new PublicClientApplication({ auth: authConfig, cache: { cachePlugin } });
}

/**
 * Returns a valid Microsoft Graph access token.
 *
 * Strategy:
 *  1. Try silent MSAL refresh using the cached refresh token
 *  2. Fall back to stored access token if it hasn't expired
 *  3. Throw if both are expired -- user must re-run auth.mjs
 *
 * @param {'personal'|'work'|'email'} account
 *   'personal' -- personal Microsoft account
 *   'work'     -- work/school single-tenant account (Teams, Planner)
 *   'email'    -- prefers personal if configured, else work (default)
 * @returns {Promise<string>} A valid access token
 */
export async function getToken(account = 'email') {
  const config = loadConfig();
  const auth = config.auth;
  const resolved = resolveAccountType(auth, account);
  const authConfig = resolveAuthConfig(auth, resolved);
  const { tokensPath, msalCachePath } = resolveTokenPaths(resolved);

  if (!existsSync(tokensPath)) {
    throw new Error(
      `No tokens found for ${resolved} account at ${tokensPath}.\n` +
      `Run: node scripts/auth.mjs --account=${resolved}`
    );
  }

  const tokens = JSON.parse(readFileSync(tokensPath, 'utf8'));

  // Personal accounts do not support Teams/Planner scopes; work accounts get
  // the full set including ChannelMessage and Tasks.
  const PERSONAL_SCOPES = [
    'https://graph.microsoft.com/Mail.Read',
    'https://graph.microsoft.com/Mail.ReadWrite',
    'https://graph.microsoft.com/Mail.Send',
    'https://graph.microsoft.com/Calendars.Read',
    'https://graph.microsoft.com/Calendars.ReadWrite',
    'https://graph.microsoft.com/offline_access',
    'https://graph.microsoft.com/User.Read',
  ];
  const WORK_SCOPES = [
    ...PERSONAL_SCOPES,
    'https://graph.microsoft.com/ChannelMessage.Read.All',
    'https://graph.microsoft.com/ChannelMessage.Send',
    'https://graph.microsoft.com/Tasks.Read',
    'https://graph.microsoft.com/Tasks.ReadWrite',
  ];
  const scopes = resolved === 'work' ? WORK_SCOPES : PERSONAL_SCOPES;

  // Attempt silent MSAL refresh
  let serializedCache = null;
  if (existsSync(msalCachePath)) {
    try { serializedCache = readFileSync(msalCachePath, 'utf8'); } catch (_) {}
  }

  if (serializedCache) {
    try {
      const pca = buildMsalApp(authConfig, serializedCache);
      const accounts = await pca.getTokenCache().getAllAccounts();
      if (accounts.length > 0) {
        const result = await pca.acquireTokenSilent({ scopes, account: accounts[0] });
        if (result?.accessToken) {
          writeFileSync(msalCachePath, pca.getTokenCache().serialize(), 'utf8');
          return result.accessToken;
        }
      }
    } catch (err) {
      if (process.env.DEBUG) console.error(`[token:${resolved}] Silent refresh failed:`, err.message);
    }
  }

  // Fall back to stored access token
  const { accessToken, expiresAt } = tokens;
  if (!accessToken) {
    throw new Error(
      `No access token in ${tokensPath}. Run: node scripts/auth.mjs --account=${resolved}`
    );
  }
  const expiry = expiresAt ? new Date(expiresAt).getTime() : 0;
  if (expiry - Date.now() < 60_000) {
    throw new Error(
      `Access token for ${resolved} account is expired and silent refresh failed.\n` +
      `Run: node scripts/auth.mjs --account=${resolved}`
    );
  }
  return accessToken;
}

/**
 * Returns the list of account types configured in config.json.
 * @returns {string[]} e.g. ['personal', 'work']
 */
export function getConfiguredAccounts() {
  const config = loadConfig();
  const auth = config.auth;
  const accounts = [];
  if (auth?.clientId) accounts.push('personal'); // legacy flat config
  if (auth?.personal?.clientId) accounts.push('personal');
  if (auth?.work?.clientId) accounts.push('work');
  return [...new Set(accounts)];
}
