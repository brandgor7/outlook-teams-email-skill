/**
 * run-todos.mjs — To-Do Creation Pipeline
 *
 * Reads the latest email summary from outlook-summary.md, uses the LLM to
 * extract action items, and creates them as tasks in Microsoft Planner.
 *
 * Usage:
 *   node scripts/run-todos.mjs
 *   node scripts/run-todos.mjs --dry-run   # print tasks, don't create
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
  if (!token) throw new Error('No gateway token in openclaw.json — ensure gateway.auth.mode is "token".');
  return token;
}

// ─── LLM call ─────────────────────────────────────────────────────────────────

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

function buildExtractionPrompt(summaryText) {
  return `Extract action items from this email summary.
Return a JSON array only. No prose. No markdown.

Each item must have:
  - "title": short task title (max 80 chars)
  - "notes": one sentence of context
  - "dueDate": ISO date if mentioned, otherwise null
  - "urgent": true/false

Summary:
${summaryText}`;
}

function parseTasks(llmOutput) {
  // Strip potential markdown code fences
  const cleaned = llmOutput
    .replace(/^```json\s*/i, '')
    .replace(/^```\s*/i, '')
    .replace(/\s*```$/, '')
    .trim();
  try {
    const parsed = JSON.parse(cleaned);
    if (!Array.isArray(parsed)) throw new Error('Expected a JSON array');
    return parsed;
  } catch (err) {
    throw new Error(`Failed to parse LLM task output: ${err.message}\n\nRaw output:\n${llmOutput}`);
  }
}

// ─── Planner task creation ────────────────────────────────────────────────────

async function createPlannerTask(token, task, config) {
  const { planId, bucketId } = config.todos.delivery;

  // Get current user id
  const meRes = await fetch(`${GRAPH}/me?$select=id`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  const me = await meRes.json();
  const userId = me?.id;

  const body = {
    title: task.title,
    planId,
    assignments: userId
      ? { [userId]: { '@odata.type': '#microsoft.graph.plannerAssignment', orderHint: ' !' } }
      : undefined,
  };
  if (bucketId) body.bucketId = bucketId;
  if (task.dueDate) body.dueDateTime = new Date(task.dueDate).toISOString();

  const res = await fetch(`${GRAPH}/planner/tasks`, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify(body),
  });

  if (!res.ok) {
    const t = await res.text().catch(() => '');
    throw new Error(`Planner API error ${res.status}: ${t}`);
  }

  const created = await res.json();

  // Add notes via task details
  if (task.notes && created?.id) {
    try {
      const detailsRes = await fetch(`${GRAPH}/planner/tasks/${created.id}/details`, {
        headers: { Authorization: `Bearer ${token}` },
      });
      const details = await detailsRes.json();
      const etag = details?.['@odata.etag'];
      if (etag) {
        await fetch(`${GRAPH}/planner/tasks/${created.id}/details`, {
          method: 'PATCH',
          headers: {
            Authorization: `Bearer ${token}`,
            'Content-Type': 'application/json',
            'If-Match': etag,
          },
          body: JSON.stringify({ description: task.notes }),
        });
      }
    } catch (_) { /* notes are optional */ }
  }

  return created?.id;
}

// ─── Main pipeline ────────────────────────────────────────────────────────────

async function main() {
  console.log(`\n✅ To-Do Creation Pipeline${DRY_RUN ? ' (DRY RUN)' : ''}\n`);

  const config = loadConfig();

  // Step 1: Check enabled
  if (!config.todos?.enabled) {
    console.log('todos.enabled is false in config.json — exiting.');
    return;
  }

  // Step 2: Read summary
  if (!existsSync(SUMMARY_PATH)) {
    console.error(`Summary file not found at ${SUMMARY_PATH}.`);
    console.error('Run run-summary.mjs first to generate a summary.');
    process.exit(1);
  }

  const summaryText = readFileSync(SUMMARY_PATH, 'utf8');
  console.log(`📄 Loaded summary from: ${SUMMARY_PATH}`);

  // Step 3: Extract action items via LLM
  console.log('🤖 Calling LLM to extract action items...');
  const prompt = buildExtractionPrompt(summaryText);
  const llmOutput = await callLlm(config, [{ role: 'user', content: prompt }]);

  let tasks;
  try {
    tasks = parseTasks(llmOutput);
  } catch (err) {
    console.error(err.message);
    process.exit(1);
  }

  console.log(`   Extracted ${tasks.length} action item(s).`);

  if (tasks.length === 0) {
    console.log('No action items found. Exiting.');
    return;
  }

  // Step 4: Dry run — print and exit
  if (DRY_RUN) {
    console.log('\n─── DRY RUN: Extracted Tasks ──────────────────────────────\n');
    tasks.forEach((t, i) => {
      const urgentFlag = t.urgent ? ' ⚠️ URGENT' : '';
      console.log(`${i + 1}. ${t.title}${urgentFlag}`);
      if (t.notes) console.log(`   Notes: ${t.notes}`);
      if (t.dueDate) console.log(`   Due:   ${t.dueDate}`);
      console.log();
    });
    console.log('───────────────────────────────────────────────────────────\n');
    return;
  }

  // Step 5: Validate Planner config
  const { planId } = config.todos.delivery;
  if (!planId) {
    console.error('todos.delivery.planId is not set in config.json.');
    console.error('Run Teams provisioning to set up Planner, or use --dry-run to test.');
    process.exit(1);
  }

  // Step 6: Get work Graph token (Planner requires work/tenant account)
  let graphToken;
  try {
    graphToken = await getToken('work');
  } catch (err) {
    console.error(`Auth error: ${err.message}`);
    process.exit(1);
  }

  // Step 7: Create tasks, skip duplicates
  const state = loadState();
  const createdTaskIds = state.createdTaskIds ?? [];

  console.log('📋 Creating Planner tasks...');
  const newIds = [];
  for (const task of tasks) {
    try {
      const id = await createPlannerTask(graphToken, task, config);
      if (id) {
        newIds.push(id);
        console.log(`   ✅ Created: ${task.title} (${id})`);
      }
    } catch (err) {
      console.error(`   ❌ Failed to create "${task.title}": ${err.message}`);
    }
  }

  // Step 8: Save state
  saveState({
    lastTodosAt: new Date().toISOString(),
    createdTaskIds: [...createdTaskIds, ...newIds],
  });

  console.log(`\n✅ To-do pipeline complete. Created ${newIds.length} task(s).\n`);
}

main().catch((err) => {
  console.error('Pipeline error:', err);
  process.exit(1);
});
