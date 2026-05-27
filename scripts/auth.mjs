/**
 * auth.mjs — OAuth Device Code Flow
 *
 * Registers the user's Microsoft account with MSAL using the Device Code flow.
 * Writes access token + expiry to outlook-tokens.json and the full MSAL cache
 * (including refresh token) to outlook-msal-cache.json.
 *
 * Usage:
 *   node scripts/auth.mjs
 */

import { readFileSync, writeFileSync, existsSync } from 'fs';
import { resolve, dirname } from 'path';
import { fileURLToPath } from 'url';
import { PublicClientApplication } from '@azure/msal-node';

const __dirname = dirname(fileURLToPath(import.meta.url));
const SKILL_ROOT = resolve(__dirname, '..');
const WORKSPACE_ROOT = resolve(SKILL_ROOT, '..', '..');

// Write tokens to workspace root if it looks like a deployed OpenClaw environment,
// otherwise write to the skill root.
function resolveOutputPath(filename) {
  const wsParent = resolve(WORKSPACE_ROOT, 'skills');
  if (existsSync(wsParent)) {
    return resolve(WORKSPACE_ROOT, filename);
  }
  return resolve(SKILL_ROOT, filename);
}

const TOKENS_PATH = resolveOutputPath('outlook-tokens.json');
const MSAL_CACHE_PATH = resolveOutputPath('outlook-msal-cache.json');
const CONFIG_PATH = resolve(SKILL_ROOT, 'config.json');

function loadConfig() {
  if (!existsSync(CONFIG_PATH)) {
    console.error(`ERROR: config.json not found at ${CONFIG_PATH}`);
    console.error('Copy config.json and set your Azure clientId first.');
    process.exit(1);
  }
  return JSON.parse(readFileSync(CONFIG_PATH, 'utf8'));
}

async function main() {
  const config = loadConfig();

  if (!config.auth?.clientId || config.auth.clientId === 'YOUR_CLIENT_ID_HERE') {
    console.error('ERROR: config.json has a placeholder clientId.');
    console.error('Set auth.clientId to your Azure app Client ID and try again.');
    process.exit(1);
  }

  let currentCache = null;
  if (existsSync(MSAL_CACHE_PATH)) {
    try {
      currentCache = readFileSync(MSAL_CACHE_PATH, 'utf8');
    } catch (_) { /* ignore */ }
  }

  const cachePlugin = {
    beforeCacheAccess: async (cacheContext) => {
      if (currentCache) {
        cacheContext.tokenCache.deserialize(currentCache);
      }
    },
    afterCacheAccess: async (cacheContext) => {
      if (cacheContext.cacheHasChanged) {
        currentCache = cacheContext.tokenCache.serialize();
      }
    },
  };

  const pca = new PublicClientApplication({
    auth: {
      clientId: config.auth.clientId,
      authority: config.auth.authority,
    },
    cache: { cachePlugin },
  });

  const scopes = [
    'Mail.Read',
    'Mail.ReadWrite',
    'Mail.Send',
    'Calendars.Read',
    'Calendars.ReadWrite',
    'ChannelMessage.Read.All',
    'ChannelMessage.Send',
    'Tasks.Read',
    'Tasks.ReadWrite',
    'offline_access',
    'User.Read',
  ];

  console.log('\n🔐 Starting Microsoft OAuth Device Code Flow...\n');

  let result;
  try {
    result = await pca.acquireTokenByDeviceCode({
      scopes,
      deviceCodeCallback: (response) => {
        console.log(response.message);
        console.log();
      },
    });
  } catch (err) {
    console.error('Authentication failed:', err.message);
    process.exit(1);
  }

  if (!result?.accessToken) {
    console.error('Authentication failed: no access token received.');
    process.exit(1);
  }

  // Write access token + expiry
  const tokenData = {
    accessToken: result.accessToken,
    expiresAt: result.expiresOn?.toISOString() ?? null,
    account: {
      username: result.account?.username,
      name: result.account?.name,
      tenantId: result.account?.tenantId,
    },
  };
  writeFileSync(TOKENS_PATH, JSON.stringify(tokenData, null, 2), 'utf8');
  console.log(`✅ Access token written to: ${TOKENS_PATH}`);

  // Write MSAL cache (contains refresh token for silent renewal)
  if (currentCache) {
    writeFileSync(MSAL_CACHE_PATH, currentCache, 'utf8');
    console.log(`✅ MSAL cache written to:   ${MSAL_CACHE_PATH}`);
  }

  console.log(`\n👤 Authenticated as: ${result.account?.username ?? 'unknown'}`);
  console.log('\nAuth complete. You can now use the MCP server and scripts.\n');
}

main().catch((err) => {
  console.error('Unexpected error:', err);
  process.exit(1);
});
