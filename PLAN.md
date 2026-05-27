# Outlook + Teams Email Assistant — Implementation Plan

A conversational email assistant built on OpenClaw with structured workflows for
summaries and to-dos. Conversational interactions (follow-up questions, drafting
replies) are handled by the agent naturally. Scheduled/triggered workflows
(summary, to-dos) run as deterministic scripts with one focused LLM call.

---

## Architecture Overview

```
User (Telegram / Teams)
        │
        ▼
   OpenClaw Agent  ◄──── Skill (SKILL.md) — instruction layer
        │
        ├── Conversational mode (ad-hoc questions, drafting)
        │         └── MCP Server tools (list_emails, get_email, etc.)
        │
        └── Structured mode (summary, to-dos)
                  └── Scripts (run-summary.mjs, run-todos.mjs)
                            ├── Microsoft Graph API (Outlook + Teams)
                            └── LLM API (one focused call per workflow)
```

---

## Directory Structure

```
~/.openclaw/workspace/
├── PLAN.md                          ← this file
├── outlook-tokens.json              ← OAuth tokens (git-ignored)
├── outlook-msal-cache.json          ← MSAL token cache (git-ignored)
├── outlook-summary.md               ← latest summary output (git-ignored)
├── run-state.json                   ← pipeline run state (git-ignored)
│
└── skills/outlook-email/            ← skill root AND Node.js package
    ├── SKILL.md                     ← instruction layer
    ├── package.json                 ← npm package (node_modules lives here)
    ├── node_modules/                ← @azure/msal-node, @modelcontextprotocol/sdk
    ├── config.json                  ← user-editable preferences
    └── scripts/
        ├── token.mjs                ← shared: MSAL token + gateway token helpers
        ├── auth.mjs                 ← OAuth provisioning (Outlook + Teams)
        ├── mcp-server.mjs           ← MCP stdio server (Outlook + Teams tools)
        ├── run-summary.mjs          ← structured summary pipeline
        └── run-todos.mjs            ← structured to-do pipeline
```

---

## Phase 1 — OAuth & Token Management ✅

### 1.1 Azure App Registration (manual, one-time)

Register a single Azure app that covers both Outlook and Teams (same Graph API):

1. Go to https://portal.azure.com → App registrations → New registration
2. Name: `OpenClaw Assistant`, Supported accounts: Personal Microsoft accounts only
3. Authentication → Add platform → Mobile and desktop → tick nativeclient URL
4. Set **Allow public client flows** → Yes
5. API Permissions → Microsoft Graph → Delegated → add all scopes:
   - `Mail.Read`
   - `Mail.ReadWrite`
   - `Mail.Send`
   - `Calendars.Read`
   - `Calendars.ReadWrite`
   - `ChannelMessage.Read.All`
   - `ChannelMessage.Send`
   - `Tasks.Read`
   - `Tasks.ReadWrite`
   - `offline_access`
   - `User.Read`
6. Grant admin consent
7. Note the **Client ID** — paste into `auth.mjs`

### 1.2 `scripts/auth.mjs` — OAuth Device Code Flow

- Uses `@azure/msal-node` PublicClientApplication
- Device code flow (user visits URL, enters short code)
- Writes tokens to `outlook-tokens.json`
- Writes MSAL cache to `outlook-msal-cache.json` for silent refresh
- Single auth covers both Outlook and Teams (same token, different Graph endpoints)

**Run once to authenticate:**
```bash
node skills/outlook-email/scripts/auth.mjs
```

### 1.3 Token Helper Module (`scripts/token.mjs`)

Shared module imported by all other scripts:
- Loads `outlook-tokens.json` + `outlook-msal-cache.json`
- Attempts silent refresh via MSAL cache
- Falls back to stored access token if not yet expired
- Throws a clear error if token is expired (instructs user to re-run auth.mjs)
- Exports a single `getToken()` async function

---

## Phase 2 — MCP Server ✅

### 2.1 `scripts/mcp-server.mjs`

Stdio MCP server registered in OpenClaw as `"outlook"`. Used by the agent for
conversational interactions. All tools call Microsoft Graph directly.

**Tools to implement:**

#### Email tools
| Tool | Args | Description |
|---|---|---|
| `list_emails` | `top`, `unread_only`, `folder` | List recent emails with subject, sender, preview, id |
| `get_email` | `id` | Fetch full body of one email by id |
| `send_email` | `to`, `subject`, `body`, `cc?` | Send an email |
| `move_email` | `id`, `folder` | Move email to folder |
| `mark_read` | `id`, `read` | Mark email read/unread |

#### Calendar tools
| Tool | Args | Description |
|---|---|---|
| `list_calendar` | `days`, `top` | List upcoming events |
| `create_event` | `subject`, `start`, `end`, `body?`, `location?`, `attendees?` | Create calendar event |

#### Teams tools
| Tool | Args | Description |
|---|---|---|
| `list_teams` | — | List Teams the user belongs to |
| `list_channels` | `teamId` | List channels in a team |
| `post_teams_message` | `teamId`, `channelId`, `message` | Post a message to a Teams channel |
| `list_tasks` | `planId?` | List tasks from Planner |
| `create_task` | `title`, `planId`, `bucketId?`, `dueDate?`, `notes?` | Create a Planner task |

**Implementation notes:**
- Import `getToken()` from `token.mjs`
- All Graph calls use `https://graph.microsoft.com/v1.0/...`
- Return clean structured JSON (no HTML, stripped body content)
- Return clear error messages on auth failure — never crash the server
- Use `@modelcontextprotocol/sdk` StdioServerTransport

**Register with OpenClaw after writing:**
```bash
openclaw mcp set outlook '{
  "command": "node",
  "args": ["/home/node/.openclaw/workspace/skills/outlook-email/scripts/mcp-server.mjs"]
}'
```

---

## Phase 3 — User Config ✅

### 3.1 `config.json` — User Preferences

All structured workflow behaviour is driven by this file. Edit to customise.

```json
{
  "summary": {
    "wordCount": 150,
    "tone": "concise and professional",
    "categories": ["urgent", "action required", "newsletters", "work", "personal", "other"],
    "maxEmails": 30,
    "unreadOnly": false,
    "delivery": {
      "channels": [
        { "type": "telegram", "chatId": "8524887570" },
        { "type": "teams", "teamId": "", "channelId": "" }
      ]
    }
  },
  "todos": {
    "enabled": true,
    "source": "summary",
    "delivery": {
      "type": "planner",
      "planId": "",
      "bucketId": ""
    }
  },
  "llm": {
    "baseUrl": "http://127.0.0.1:18789/v1",
    "model": "openclaw/default",
    "gatewayConfigPath": "~/.openclaw/openclaw.json",
    "modelOverride": ""
  },
  "auth": {
    "clientId": "YOUR_CLIENT_ID_HERE",
    "authority": "https://login.microsoftonline.com/consumers"
  }
}
```

**Notes:**
- `delivery.channels` is an array — deliver to multiple channels simultaneously
- Teams channel delivery requires `teamId` + `channelId` (get these from `list_teams` / `list_channels` MCP tools)
- `llm.baseUrl` + `llm.model` — points at the local OpenClaw gateway's OpenAI-compatible endpoint; no separate API key needed
- `llm.gatewayConfigPath` — path to `openclaw.json`; scripts read `gateway.auth.token` from it at runtime (no env var, no hardcoding)
- `llm.modelOverride` — optional: pin a specific backend model (e.g. `openai/gpt-4o-mini`) for summarisation independent of your main chat model; leave blank to use whatever OpenClaw has active
- Leave `planId`/`bucketId` blank initially — fill after running Teams provisioning

---

## Phase 4 — Structured Workflows ✅

### 4.1 `scripts/run-summary.mjs` — Email Summary Pipeline

Deterministic pipeline. LLM is called once with a fixed prompt template.

**Pipeline:**
```
1. Load config.json
2. getToken()
3. Fetch emails from Graph API (respects config.maxEmails, config.unreadOnly)
4. Filter out already-processed emails (optional: track lastRunAt timestamp)
5. Build LLM prompt from fixed template (inject config.tone, config.wordCount, config.categories)
6. Call LLM API → get summary prose
7. Inject prose into output template (fixed markdown structure)
8. Deliver to each configured channel:
   a. Telegram → Bot API sendMessage
   b. Teams → Graph API post to channel
9. Write summary to outlook-summary.md (for conversational follow-up)
10. Write lastRunAt to run-state.json
```

**LLM prompt template:**
```
You are summarising emails for a busy professional.

Summarise the following {{emailCount}} emails in approximately {{wordCount}} words.
Tone: {{tone}}
Group emails into these categories (omit empty categories): {{categories}}
For each email include: sender, subject, one-sentence summary.
Flag anything marked urgent or requiring a reply with ⚠️.

Return ONLY the summary text. No preamble. No JSON.

Emails:
{{emailsJSON}}
```

**Output template (fixed structure, LLM only fills the prose block):**
```markdown
## 📬 Email Summary — {{datetime}}
_{{unreadCount}} unread · {{totalFetched}} checked · via OpenClaw_

{{llmSummaryProse}}

---
_Reply "check my emails" to ask about any of these._
```

**CLI usage:**
```bash
node skills/outlook-email/scripts/run-summary.mjs
node skills/outlook-email/scripts/run-summary.mjs --dry-run   # skip delivery, print to stdout
```

### 4.2 `scripts/run-todos.mjs` — To-Do Creation Pipeline

Reads the latest summary and extracts action items into Planner tasks.

**Pipeline:**
```
1. Load config.json
2. Check todos.enabled — exit cleanly if false
3. Read outlook-summary.md (written by run-summary.mjs)
4. Call LLM API with fixed extraction prompt → get structured JSON list of action items
5. For each action item:
   a. POST to Microsoft Planner (Graph API) → create task
   b. Log created task id to run-state.json (avoid duplicates)
6. Post confirmation message to configured Teams channel (optional)
```

**LLM extraction prompt template:**
```
Extract action items from this email summary.
Return a JSON array only. No prose. No markdown.

Each item must have:
  - "title": short task title (max 80 chars)
  - "notes": one sentence of context
  - "dueDate": ISO date if mentioned, otherwise null
  - "urgent": true/false

Summary:
{{summaryText}}
```

**CLI usage:**
```bash
node skills/outlook-email/scripts/run-todos.mjs
node skills/outlook-email/scripts/run-todos.mjs --dry-run   # print tasks, don't create
```

### 4.3 `run-state.json` (auto-generated, git-ignored)

Tracks pipeline state between runs:
```json
{
  "lastSummaryAt": "2026-05-26T20:00:00Z",
  "lastTodosAt": "2026-05-26T20:01:00Z",
  "createdTaskIds": ["task-id-1", "task-id-2"]
}
```

---

## Phase 5 — Skill (Instruction Layer) ✅

### 5.1 `SKILL.md`

Tells the agent:

**When to use MCP tools directly (conversational mode):**
- "Check my emails" → `list_emails`
- "Read that email from John" → `get_email`
- "Draft a reply to..." → `get_email` then compose
- "What's on my calendar?" → `list_calendar`
- "Create a meeting..." → `create_event`
- "Post to Teams..." → `post_teams_message`

**When to call scripts (structured mode):**
- "Run the summary" / scheduled cron → `run-summary.mjs`
- "Create to-dos from my emails" → `run-todos.mjs`
- Always pass `--dry-run` first if user wants to preview

**Config awareness:**
- Before structured workflows, read `config.json` — mention active settings to the user
- If config values are missing (empty teamId, planId), prompt user to run provisioning

**Auth failure handling:**
- If any MCP tool returns token expired → tell user to run `auth.mjs`
- Never silently swallow auth errors

**Delivery:**
- Structured workflow output goes to channels defined in `config.json`
- Conversational output stays in the current chat session

---

## Phase 6 — Cron Jobs ✅

Set up after everything is tested manually.

### 6.1 Summary (every 6 hours)
```bash
openclaw cron add \
  --name "Outlook Email Summary" \
  --every "6h" \
  --session isolated \
  --message "Run the email summary workflow by executing: node /home/node/.openclaw/workspace/skills/outlook-email/scripts/run-summary.mjs — then confirm delivery completed or report any errors."
```

### 6.2 To-Dos (after each summary)
```bash
openclaw cron add \
  --name "Email To-Dos" \
  --every "6h" \
  --session isolated \
  --message "Run the to-do creation workflow by executing: node /home/node/.openclaw/workspace/skills/outlook-email/scripts/run-todos.mjs — then confirm tasks were created or report any errors."
```

Offset the to-do job by ~5 minutes so the summary always runs first.

---

## Phase 7 — Provisioning (Setup Flows) ✅

Guided setup conversations using `auth.mjs` + the MCP tools to discover IDs.

### 7.1 Outlook Setup
1. User says "set up Outlook" → agent guides them through Azure app registration steps
2. User pastes Client ID → agent writes it to `config.json`
3. Agent runs `auth.mjs` (user completes device code flow in browser)
4. Agent calls `list_emails` to verify connection

### 7.2 Teams Setup
1. User says "set up Teams" → agent calls `list_teams` MCP tool
2. Presents list of teams → user picks one
3. Agent calls `list_channels` → user picks delivery channel
4. Agent writes `teamId` + `channelId` to `config.json`
5. Agent calls `list_tasks` to find Planner plans → user picks one
6. Agent writes `planId` to `config.json`
7. Agent runs `run-summary.mjs --dry-run` to confirm output format

---

## Implementation Order

1. **`token.mjs`** — shared token helper (everything else depends on this)
2. **`auth.mjs`** — OAuth flow (needed before any Graph calls work)
3. **`mcp-server.mjs`** — MCP server with all tools, register with OpenClaw
4. **`config.json`** — fill in Client ID, Telegram chat ID; leave Teams IDs for later
5. **`run-summary.mjs`** — test with `--dry-run` before enabling delivery
6. **`run-todos.mjs`** — test with `--dry-run` before creating real tasks
7. **`SKILL.md`** — write instruction layer once scripts are working
8. **Cron jobs** — add after manual testing passes
9. **Teams provisioning** — run setup conversation to fill in Teams config IDs

---

## Dependencies

The skill root (`skills/outlook-email/`) doubles as the Node.js package — `package.json` and `node_modules` live there:

```bash
cd ~/.openclaw/workspace/skills/outlook-email
npm init -y
npm install @azure/msal-node @modelcontextprotocol/sdk
```

Since `node_modules` is in the skill root, all scripts import dependencies with standard ESM — no `createRequire` workaround needed:

```js
import { PublicClientApplication } from '@azure/msal-node';
import { Server } from '@modelcontextprotocol/sdk/server/index.js';
```

Node resolves `node_modules` by walking up from `scripts/` to the parent directory automatically. LLM calls, Graph API calls, and Telegram Bot API calls all use native `fetch` — no extra dependencies.

## LLM Integration — OpenClaw Gateway Proxy

Scripts call OpenClaw's OpenAI-compatible HTTP endpoint instead of an external provider directly:

```
POST http://127.0.0.1:18789/v1/chat/completions
Authorization: Bearer <gateway-token>
x-openclaw-model: <optional model override>
```

This routes through whatever provider + model OpenClaw is currently configured with — no separate API key needed in scripts.

**Gateway setup required:**
- Enable `chatCompletions` endpoint in `openclaw.json`:
  ```bash
  echo '{"gateway":{"http":{"endpoints":{"chatCompletions":{"enabled":true}}}}}' | openclaw config patch --stdin
  openclaw gateway restart
  ```

**How scripts read the token at runtime:**

The gateway token lives inline in `openclaw.json` under `gateway.auth.token`. Scripts read it directly from the config file — no env var, no hardcoding:

```js
import { readFileSync } from 'fs';
import { resolve } from 'path';
import { homedir } from 'os';

function getGatewayToken(config) {
  const configPath = resolve(config.llm.gatewayConfigPath.replace('~', homedir()));
  const gw = JSON.parse(readFileSync(configPath, 'utf8'));
  const token = gw?.gateway?.auth?.token;
  if (!token) throw new Error('No gateway token found in openclaw.json — is gateway.auth.mode set to "token"?');
  return token;
}

// Usage:
const token = getGatewayToken(config);
const headers = {
  'Authorization': `Bearer ${token}`,
  'Content-Type': 'application/json',
  ...(config.llm.modelOverride ? { 'x-openclaw-model': config.llm.modelOverride } : {}),
};
const res = await fetch(`${config.llm.baseUrl}/chat/completions`, {
  method: 'POST',
  headers,
  body: JSON.stringify({ model: config.llm.model, messages: [...] }),
});
```
