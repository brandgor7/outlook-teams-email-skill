/**
 * auth.mjs -- OAuth Device Code Flow
 *
 * Authenticates with Microsoft using the Device Code flow.
 * Supports personal Microsoft accounts and work/school (single-tenant) accounts.
 *
 * Usage:
 *   node scripts/auth.mjs --account=personal
 *   node scripts/auth.mjs --account=work
 *   node scripts/auth.mjs           # auto-detects from config.json
 *
 * npm shortcuts:
 *   npm run auth:personal
 *   npm run auth:work
 */

import { readFileSync, writeFileSync, existsSync } from 'fs';
import { resolve, dirname } from 'path';
import { fileURLToPath } from 'url';
import { PublicClientApplication } from '@azure/msal-node';

const __dirname = dirname(fileURLToPath(import.meta.url));
const SKILL_ROOT = resolve(__dirname, '..');
const WORKSPACE_ROOT = resolve(SKILL_ROOT, '..', '..');
const CONFIG_PATH = resolve(SKILL_ROOT, 'config.json');

// Write tokens to workspace root if deployed under OpenClaw, else skill root.
function resolveOutputPath(filename) {
  const wsParent = resolve(WORKSPACE_ROOT, 'skills');
  if (existsSync(wsParent)) return resolve(WORKSPACE_ROOT, filename);
  return resolve(SKILL_ROOT, filename);
}

function loadConfig() {
  if (!existsSync(CONFIG_PATH)) {
    console.error(`ERROR: config.json not found at ${CONFIG_PATH}`);
    console.error('Copy config.json and set your Azure credentials first.');
    process.exit(1);
  }
  return JSON.parse(readFileSync(CONFIG_PATH, 'utf8'));
}

// Parse --account=personal or --account=work from argv
function parseAccountArg() {
  const arg = process.argv.find((a) => a.startsWith('--account='));
  if (arg) return arg.split('=')[1];
  return null;
}

// Determine account type from config when not explicitly provided
function detectAccount(config) {
  const auth = config.auth;
  if (auth?.clientId) return 'personal'; // legacy flat config
  const hasPersonal = !!(auth?.personal?.clientId);
  const hasWork = !!(auth?.work?.clientId);
  if (hasPersonal && hasWork) {
    console.error('ERROR: Both auth.personal and auth.work are configured.');
    console.error('Specify which account to authenticate:');
    console.error('  node scripts/auth.mjs --account=personal');
    console.error('  node scripts/auth.mjs --account=work');
    process.exit(1);
  }
  if (hasPersonal) return 'personal';
  if (hasWork) return 'work';
  console.error('ERROR: No auth configuration found in config.json.');
  console.error('Configure auth.personal and/or auth.work, then re-run.');
  process.exit(1);
}

// Build MSAL auth config for the selected account
function buildAuthConfig(config, account) {
  const auth = config.auth;

  if (account === 'personal') {
    // Support legacy flat format
    const clientId = auth?.personal?.clientId ?? auth?.clientId;
    const authority = auth?.personal?.authority ?? auth?.authority ?? 'https://login.microsoftonline.com/consumers';
    if (!clientId || clientId.startsWith('YOUR_')) {
      console.error('ERROR: auth.personal.clientId is a placeholder or missing in config.json.');
      console.error('Set it to your Azure app Client ID and try again.');
      process.exit(1);
    }
    return { clientId, authority };
  }

  if (account === 'work') {
    const clientId = auth?.work?.clientId;
    const tenantId = auth?.work?.tenantId;
    const authority = auth?.work?.authority ?? (tenantId ? `https://login.microsoftonline.com/${tenantId}` : null);
    if (!clientId || clientId.startsWith('YOUR_')) {
      console.error('ERROR: auth.work.clientId is a placeholder or missing in config.json.');
      console.error('Set it to your Azure work app Client ID and try again.');
      process.exit(1);
    }
    if (!authority) {
      console.error('ERROR: auth.work.tenantId (or auth.work.authority) is missing in config.json.');
      process.exit(1);
    }
    return { clientId, authority };
  }

  console.error(`ERROR: Unknown account type: "${account}". Use 'personal' or 'work'.`);
  process.exit(1);
}

async function main() {
  const config = loadConfig();
  const accountArg = parseAccountArg();
  const account = accountArg ?? detectAccount(config);
  const authConfig = buildAuthConfig(config, account);

  const tokensFile = account === 'work' ? 'work-tokens.json' : 'personal-tokens.json';
  const cacheFile = account === 'work' ? 'work-msal-cache.json' : 'personal-msal-cache.json';
  const TOKENS_PATH = resolveOutputPath(tokensFile);
  const MSAL_CACHE_PATH = resolveOutputPath(cacheFile);

  let currentCache = null;
  if (existsSync(MSAL_CACHE_PATH)) {
    try { currentCache = readFileSync(MSAL_CACHE_PATH, 'utf8'); } catch (_) {}
  }

  const cachePlugin = {
    beforeCacheAccess: async (cacheContext) => {
      if (currentCache) cacheContext.tokenCache.deserialize(currentCache);
    },
    afterCacheAccess: async (cacheContext) => {
      if (cacheContext.cacheHasChanged) {
        currentCache = cacheContext.tokenCache.serialize();
      }
    },
  };

  const pca = new PublicClientApplication({
    auth: authConfig,
    cache: { cachePlugin },
  });

  // Personal accounts do not support Teams/Planner scopes; work accounts get
  // the full set including ChannelMessage and Tasks.
  const PERSONAL_SCOPES = [
    'Mail.Read',
    'Mail.ReadWrite',
    'Mail.Send',
    'Calendars.Read',
    'Calendars.ReadWrite',
    'offline_access',
    'User.Read',
  ];
  const WORK_SCOPES = [
    'ChannelMessage.Read.All',
    'ChannelMessage.Send',
    'Tasks.Read',
    'Tasks.ReadWrite',
    'offline_access',
    'User.Read',
  ];
  const scopes = account === 'work' ? WORK_SCOPES : PERSONAL_SCOPES;

  console.log(`\n🔐 Starting Microsoft OAuth Device Code Flow (${account} account)...\n`);

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
