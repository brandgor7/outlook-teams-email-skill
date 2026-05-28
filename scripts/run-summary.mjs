/**
 * run-summary.mjs — Email Summary Pipeline
 *
 * Fetches recent emails from Microsoft Graph, summarises them using the
 * OpenClaw LLM gateway, and delivers the output to configured channels
 * (Telegram and/or Teams).
 *
 * Usage:
 *   node scripts/run-summary.mjs
 *   node scripts/run-summary.mjs --dry-run   # print to stdout, skip delivery
 */

import { readFileSync, writeFileSync, existsSync } from 'fs';
import { resolve, dirname } from 'path';
import { fileURLToPath } from 'url';
import { homedir } from 'os';
import { getToken } from './token.mjs';

const __dirname = dirname(fileURLToPath(import.meta.url));
const SKILL_ROOT = resolve(__dirname, '..');
const WORKSPACE_ROOT = resolve(SKILL_ROOT, '..', '..');

const DRY_RUN = process.argv.includes('--dry-run');
const GRAPH = 'https://graph.microsoft.com/v1.0';

// ─── Path helpers ─────────────────────────────────────────────────────────────

function resolveStatePath(filename) {
  const wsParent = resolve(WORKSPACE_ROOT, 'skills');
  if (existsSync(wsParent)) return resolve(WORKSPACE_ROOT, filename);
  return resolve(SKILL_ROOT, filename);
}

const CONFIG_PATH = resolve(SKILL_ROOT, 'config.json');
const SUMMARY_PATH = resolveStatePath('outlook-summary.md');
const STATE_PATH = resolveStatePath('run-state.json');

// ─── Config & state ───────────────────────────────────────────────────────────

function loadConfig() {
  return JSON.parse(readFileSync(CONFIG_PATH, 'utf8'));
}

function loadState() {
  if (!existsSync(STATE_PATH)) return {};
  try { return JSON.parse(readFileSync(STATE_PATH, 'utf8')); } catch (_) { return {}; }
}

function saveState(state) {
  const current = loadState();
  writeFileSync(STATE_PATH, JSON.stringify({ ...current, ...state }, null, 2), 'utf8');
}

// ─── LLM gateway token ────────────────────────────────────────────────────────

function getGatewayToken(config) {
  const configPath = resolve(config.llm.gatewayConfigPath.replace('~', homedir()));
  if (!existsSync(configPath)) {
    throw new Error(`openclaw.json not found at ${configPath}. Is OpenClaw installed?`);
  }
  const gw = JSON.parse(readFileSync(configPath, 'utf8'));
  const token = gw?.gateway?.auth?.token;
  if (!token) throw new Error('No gateway token found in openclaw.json — ensure gateway.auth.mode is "token".');
  return token;
}

// ─── Graph API helpers ────────────────────────────────────────────────────────

async function graphGet(token, path) {
  const res = await fetch(`${GRAPH}${path}`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!res.ok) {
    const t = await res.text().catch(() => '');
    throw new Error(`Graph ${res.status}: ${res.statusText} ${t}`);
  }
  return res.json();
}

function stripHtml(html) {
  if (!html) return '';
  return html
    .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, '')
    .replace(/<script[^>]*>[\s\S]*?<\/script>/gi, '')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/\s+/g, ' ')
    .trim();
}

// ─── Email fetching ───────────────────────────────────────────────────────────

async function fetchEmails(token, config) {
  const { maxEmails = 30, unreadOnly = false } = config.summary;
  const filter = unreadOnly ? '&$filter=isRead eq false' : '';
  const select = '$select=id,subject,from,receivedDateTime,isRead,body,bodyPreview';
  const data = await graphGet(
    token,
    `/me/mailFolders/inbox/messages?$top=${maxEmails}&${select}&$orderby=receivedDateTime desc${filter}`
  );
  return (data?.value ?? []).map((m) => ({
    id: m.id,
    subject: m.subject ?? '(no subject)',
    from: m.from?.emailAddress?.address ?? 'unknown',
    fromName: m.from?.emailAddress?.name ?? '',
    receivedAt: m.receivedDateTime,
    isRead: m.isRead,
    body: m.body?.contentType === 'html' ? stripHtml(m.body.content) : (m.body?.content ?? m.bodyPreview ?? ''),
  }));
}

// ─── LLM summarisation ────────────────────────────────────────────────────────

async function callLlm(config, messages) {
  const gatewayToken = getGatewayToken(config);
  const headers = {
    Authorization: `Bearer ${gatewayToken}`,
    'Content-Type': 'application/json',
    ...(config.llm.modelOverride ? { 'x-openclaw-model': config.llm.modelOverride } : {}),
  };
  const res = await fetch(`${config.llm.baseUrl}/chat/completions`, {
    method: 'POST',
    headers,
    body: JSON.stringify({ model: config.llm.model, messages }),
  });
  if (!res.ok) {
    const t = await res.text().catch(() => '');
    throw new Error(`LLM gateway error ${res.status}: ${t}`);
  }
  const data = await res.json();
  return data?.choices?.[0]?.message?.content ?? '';
}

function buildSummaryPrompt(emails, config) {
  const { wordCount = 150, tone = 'concise and professional', categories = [] } = config.summary;
  const emailsJSON = JSON.stringify(
    emails.map((e) => ({
      from: e.fromName ? `${e.fromName} <${e.from}>` : e.from,
      subject: e.subject,
      receivedAt: e.receivedAt,
      body: e.body.slice(0, 500), // trim to keep prompt manageable
    })),
    null, 2
  );
  return `You are summarising emails for a busy professional.

Summarise the following ${emails.length} emails in approximately ${wordCount} words.
Tone: ${tone}
Group emails into these categories (omit empty categories): ${categories.join(', ')}
For each email include: sender, subject, one-sentence summary.
Flag anything marked urgent or requiring a reply with ⚠️.

Return ONLY the summary text. No preamble. No JSON.

Emails:
${emailsJSON}`;
}

// ─── Delivery ─────────────────────────────────────────────────────────────────

async function deliverTelegram(channel, text) {
  const { chatId } = channel;
  const botToken = process.env.TELEGRAM_BOT_TOKEN;
  if (!botToken) {
    console.warn('TELEGRAM_BOT_TOKEN env var not set — skipping Telegram delivery.');
    return;
  }
  const res = await fetch(`https://api.telegram.org/bot${botToken}/sendMessage`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      chat_id: chatId,
      text,
      parse_mode: 'Markdown',
    }),
  });
  if (!res.ok) {
    const t = await res.text().catch(() => '');
    throw new Error(`Telegram API error ${res.status}: ${t}`);
  }
  console.log(`✅ Delivered to Telegram chat ${chatId}`);
}

async function deliverTeams(channel, text) {
  const { teamId, channelId } = channel;
  if (!teamId || !channelId) {
    console.warn('Teams teamId or channelId not configured in config.json -- skipping Teams delivery.');
    return;
  }
  let workToken;
  try {
    workToken = await getToken('work');
  } catch (err) {
    console.warn(`Teams delivery skipped: ${err.message}`);
    return;
  }
  const htmlText = text.replace(/\n/g, '<br>');
  const res = await fetch(
    `${GRAPH}/teams/${encodeURIComponent(teamId)}/channels/${encodeURIComponent(channelId)}/messages`,
    {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${workToken}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({ body: { contentType: 'html', content: htmlText } }),
    }
  );
  if (!res.ok) {
    const t = await res.text().catch(() => '');
    throw new Error(`Teams API error ${res.status}: ${t}`);
  }
  console.log(`✅ Delivered to Teams channel ${channelId}`);
}

async function deliverAll(config, summaryText) {
  const channels = config.summary?.delivery?.channels ?? [];
  for (const channel of channels) {
    try {
      if (channel.type === 'telegram') {
        await deliverTelegram(channel, summaryText);
      } else if (channel.type === 'teams') {
        await deliverTeams(channel, summaryText);
      } else {
        console.warn(`Unknown delivery channel type: ${channel.type}`);
      }
    } catch (err) {
      console.error(`Delivery to ${channel.type} failed: ${err.message}`);
    }
  }
}

// ─── Main pipeline ────────────────────────────────────────────────────────────

async function main() {
  console.log(`\n📬 Email Summary Pipeline${DRY_RUN ? ' (DRY RUN)' : ''}\n`);

  const config = loadConfig();

  // Step 1: Get Graph token for email (personal account preferred, falls back to work)
  let graphToken;
  try {
    graphToken = await getToken('email');
  } catch (err) {
    console.error(`Auth error: ${err.message}`);
    process.exit(1);
  }

  // Step 2: Fetch emails
  console.log('📥 Fetching emails from Microsoft Graph...');
  const emails = await fetchEmails(graphToken, config);
  console.log(`   Found ${emails.length} email(s).`);

  if (emails.length === 0) {
    console.log('No emails to summarise. Exiting.');
    return;
  }

  const unreadCount = emails.filter((e) => !e.isRead).length;

  // Step 3: Build prompt and call LLM
  console.log('🤖 Calling LLM for summary...');
  const prompt = buildSummaryPrompt(emails, config);
  const summaryProse = await callLlm(config, [{ role: 'user', content: prompt }]);

  // Step 4: Build output markdown
  const now = new Date();
  const datetime = now.toLocaleString('en-US', {
    weekday: 'short', year: 'numeric', month: 'short', day: 'numeric',
    hour: '2-digit', minute: '2-digit',
  });
  const summaryOutput = `## 📬 Email Summary — ${datetime}
_${unreadCount} unread · ${emails.length} checked · via OpenClaw_

${summaryProse}

---
_Reply "check my emails" to ask about any of these._`;

  // Step 5: Dry run — print and exit
  if (DRY_RUN) {
    console.log('\n─── DRY RUN OUTPUT ────────────────────────────────────────\n');
    console.log(summaryOutput);
    console.log('\n───────────────────────────────────────────────────────────\n');
    return;
  }

  // Step 6: Deliver
  console.log('📤 Delivering summary...');
  await deliverAll(config, summaryOutput);

  // Step 7: Write summary file for run-todos.mjs
  writeFileSync(SUMMARY_PATH, summaryOutput, 'utf8');
  console.log(`📝 Summary written to: ${SUMMARY_PATH}`);

  // Step 8: Update run state
  saveState({ lastSummaryAt: now.toISOString() });

  console.log('\n✅ Summary pipeline complete.\n');
}

main().catch((err) => {
  console.error('Pipeline error:', err);
  process.exit(1);
});
