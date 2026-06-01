/**
 * run-summary.mjs — Email Fetch + Delivery Pipeline
 *
 * Two modes of operation:
 *
 *   (default)   Fetch recent emails from Microsoft Graph and print them as
 *               formatted text for the OpenClaw agent to read and summarise.
 *               Also writes raw email data to outlook-emails.json.
 *               The agent uses SKILL.md "Email Summary Format" to produce the
 *               summary, then writes it to outlook-summary.md and calls
 *               `node scripts/run-summary.mjs --deliver`.
 *
 *   --deliver   Read outlook-summary.md and post it to all configured delivery
 *               channels (Telegram, Teams).  Combine with --dry-run to preview
 *               without actually sending.
 *
 * Usage:
 *   node scripts/run-summary.mjs                      # fetch + print emails
 *   node scripts/run-summary.mjs --deliver            # deliver outlook-summary.md
 *   node scripts/run-summary.mjs --deliver --dry-run  # preview delivery output
 */

import { readFileSync, writeFileSync, existsSync } from 'fs';
import { resolve, dirname } from 'path';
import { fileURLToPath } from 'url';
import { getToken } from './token.mjs';

const __dirname = dirname(fileURLToPath(import.meta.url));
const SKILL_ROOT = resolve(__dirname, '..');
const WORKSPACE_ROOT = resolve(SKILL_ROOT, '..', '..');

const DELIVER_MODE = process.argv.includes('--deliver');
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

// ─── Email paths ──────────────────────────────────────────────────────────────

const EMAILS_PATH = resolveStatePath('outlook-emails.json');

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

// --deliver: read outlook-summary.md and post to all configured channels
async function runDeliver() {
  console.log(`\n📤 Email Summary Delivery${DRY_RUN ? ' (DRY RUN)' : ''}\n`);

  if (!existsSync(SUMMARY_PATH)) {
    console.error(`No summary file found at: ${SUMMARY_PATH}`);
    console.error('Generate the summary first, then re-run with --deliver.');
    process.exit(1);
  }

  const summaryText = readFileSync(SUMMARY_PATH, 'utf8');
  const config = loadConfig();

  if (DRY_RUN) {
    console.log('\n─── DRY RUN — would deliver ───────────────────────────────\n');
    console.log(summaryText);
    console.log('\n───────────────────────────────────────────────────────────\n');
    return;
  }

  await deliverAll(config, summaryText);
  saveState({ lastDeliveredAt: new Date().toISOString() });
  console.log('\n✅ Delivery complete.\n');
}

// default: fetch emails and print them for the OpenClaw agent to summarise
async function runFetch() {
  console.log('\n📬 Email Fetch\n');

  const config = loadConfig();

  let graphToken;
  try {
    graphToken = await getToken('email');
  } catch (err) {
    console.error(`Auth error: ${err.message}`);
    process.exit(1);
  }

  console.log('📥 Fetching emails from Microsoft Graph...');
  const emails = await fetchEmails(graphToken, config);
  console.log(`   Found ${emails.length} email(s).\n`);

  if (emails.length === 0) {
    console.log('No emails found. Nothing to summarise.');
    return;
  }

  const unreadCount = emails.filter((e) => !e.isRead).length;
  const { wordCount = 150, tone = 'concise and professional', categories = [] } = config.summary;

  // Print context header for the agent
  console.log('─── EMAIL DATA FOR SUMMARISATION ──────────────────────────');
  console.log(`Unread: ${unreadCount} / Total fetched: ${emails.length}`);
  console.log(`Requested word count: ~${wordCount} | Tone: ${tone}`);
  console.log(`Categories: ${categories.join(', ')}`);
  console.log('────────────────────────────────────────────────────────────\n');

  for (const e of emails) {
    const sender = e.fromName ? `${e.fromName} <${e.from}>` : e.from;
    const readFlag = e.isRead ? '' : ' [UNREAD]';
    console.log(`From: ${sender}${readFlag}`);
    console.log(`Subject: ${e.subject}`);
    console.log(`Received: ${e.receivedAt}`);
    console.log(`Body: ${e.body.slice(0, 400)}${e.body.length > 400 ? '...' : ''}`);
    console.log('');
  }

  console.log('────────────────────────────────────────────────────────────');
  console.log('Summarise the emails above following the "Email Summary Format"');
  console.log('defined in SKILL.md. Write the result to outlook-summary.md,');
  console.log('then run: node scripts/run-summary.mjs --deliver');
  console.log('────────────────────────────────────────────────────────────\n');

  // Also persist raw data for reference
  writeFileSync(
    EMAILS_PATH,
    JSON.stringify({ fetchedAt: new Date().toISOString(), unreadCount, emails }, null, 2),
    'utf8'
  );
  console.log(`📁 Raw email data written to: ${EMAILS_PATH}\n`);

  saveState({ lastFetchAt: new Date().toISOString() });
}

async function main() {
  if (DELIVER_MODE) {
    await runDeliver();
  } else {
    await runFetch();
  }
}

main().catch((err) => {
  console.error('Pipeline error:', err);
  process.exit(1);
});
