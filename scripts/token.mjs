/**
 * token.mjs — Shared MSAL token helper
 *
 * Loads outlook-tokens.json + outlook-msal-cache.json from the workspace root
 * (~/.openclaw/workspace/ when deployed, or two levels up from scripts/ here).
 * Attempts silent MSAL refresh first; falls back to stored access token if
 * still valid; throws a clear error if everything is expired.
 */

import { readFileSync, existsSync } from 'fs';
import { resolve, dirname } from 'path';
import { fileURLToPath } from 'url';
import { PublicClientApplication } from '@azure/msal-node';

const __dirname = dirname(fileURLToPath(import.meta.url));

// Workspace root is two directories up from scripts/ (skill root → workspace root).
// When this skill is deployed at ~/.openclaw/workspace/skills/outlook-email/,
// workspace root is ~/.openclaw/workspace/.
// During dev the skill root IS the project root, so workspaceRoot = project root.
const SKILL_ROOT = resolve(__dirname, '..');
const WORKSPACE_ROOT = resolve(SKILL_ROOT, '..', '..');

function resolveWorkspacePath(filename) {
  // Prefer workspace root (deployed location), fall back to skill root (dev)
  const wsPath = resolve(WORKSPACE_ROOT, filename);
  const skillPath = resolve(SKILL_ROOT, filename);
  if (existsSync(wsPath)) return wsPath;
  return skillPath; // may not exist yet — callers handle missing files
}

const TOKENS_PATH = resolveWorkspacePath('outlook-tokens.json');
const MSAL_CACHE_PATH = resolveWorkspacePath('outlook-msal-cache.json');
const CONFIG_PATH = resolve(SKILL_ROOT, 'config.json');

function loadConfig() {
  if (!existsSync(CONFIG_PATH)) {
    throw new Error(`config.json not found at ${CONFIG_PATH}. Copy config.json and fill in clientId.`);
  }
  return JSON.parse(readFileSync(CONFIG_PATH, 'utf8'));
}

function loadTokens() {
  if (!existsSync(TOKENS_PATH)) {
    throw new Error(
      `No tokens found at ${TOKENS_PATH}.\nRun: node scripts/auth.mjs`
    );
  }
  return JSON.parse(readFileSync(TOKENS_PATH, 'utf8'));
}

function buildMsalApp(config, serializedCache) {
  const cachePlugin = {
    beforeCacheAccess: async (cacheContext) => {
      if (serializedCache) {
        cacheContext.tokenCache.deserialize(serializedCache);
      }
    },
    afterCacheAccess: async (_cacheContext) => {
      // Cache write-back is handled in getToken() after silent refresh
    },
  };

  return new PublicClientApplication({
    auth: {
      clientId: config.auth.clientId,
      authority: config.auth.authority,
    },
    cache: { cachePlugin },
  });
}

/**
 * Returns a valid Microsoft Graph access token.
 *
 * Strategy:
 *  1. Try silent MSAL refresh using the cached refresh token
 *  2. Fall back to stored access token if it hasn't expired
 *  3. Throw if both are expired — user must re-run auth.mjs
 *
 * @returns {Promise<string>} A valid access token
 */
export async function getToken() {
  const config = loadConfig();
  const tokens = loadTokens();

  const scopes = [
    'https://graph.microsoft.com/Mail.Read',
    'https://graph.microsoft.com/Mail.ReadWrite',
    'https://graph.microsoft.com/Mail.Send',
    'https://graph.microsoft.com/Calendars.Read',
    'https://graph.microsoft.com/Calendars.ReadWrite',
    'https://graph.microsoft.com/ChannelMessage.Read.All',
    'https://graph.microsoft.com/ChannelMessage.Send',
    'https://graph.microsoft.com/Tasks.Read',
    'https://graph.microsoft.com/Tasks.ReadWrite',
    'https://graph.microsoft.com/offline_access',
    'https://graph.microsoft.com/User.Read',
  ];

  // Attempt silent refresh via MSAL cache
  let serializedCache = null;
  if (existsSync(MSAL_CACHE_PATH)) {
    try {
      serializedCache = readFileSync(MSAL_CACHE_PATH, 'utf8');
    } catch (_) { /* ignore */ }
  }

  if (serializedCache) {
    try {
      const pca = buildMsalApp(config, serializedCache);
      const accounts = await pca.getTokenCache().getAllAccounts();
      if (accounts.length > 0) {
        const result = await pca.acquireTokenSilent({
          scopes,
          account: accounts[0],
        });
        if (result?.accessToken) {
          // Write back refreshed cache
          const { writeFileSync } = await import('fs');
          writeFileSync(MSAL_CACHE_PATH, pca.getTokenCache().serialize(), 'utf8');
          return result.accessToken;
        }
      }
    } catch (err) {
      // Silent refresh failed — fall through to stored token check
      if (process.env.DEBUG) console.error('[token] Silent refresh failed:', err.message);
    }
  }

  // Fall back to stored access token
  const { accessToken, expiresAt } = tokens;
  if (!accessToken) {
    throw new Error(
      'No access token in outlook-tokens.json. Run: node scripts/auth.mjs'
    );
  }

  const now = Date.now();
  const expiry = expiresAt ? new Date(expiresAt).getTime() : 0;
  // Allow a 60-second buffer before expiry
  if (expiry - now < 60_000) {
    throw new Error(
      'Access token is expired and silent refresh failed.\nRun: node scripts/auth.mjs'
    );
  }

  return accessToken;
}
