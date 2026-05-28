# Outlook + Teams Email Assistant

An OpenClaw skill that connects to Microsoft Outlook and Teams via the Microsoft
Graph API. Provides conversational email/calendar/tasks access and scheduled
email summaries with to-do extraction.

---

## Table of Contents

1. [Prerequisites](#prerequisites)
2. [Installation](#installation)
3. [Azure App Registration](#azure-app-registration)
4. [Authentication](#authentication)
5. [MCP Server Registration](#mcp-server-registration)
6. [Configuration](#configuration)
7. [Structured Workflows](#structured-workflows)
8. [Skill Registration](#skill-registration)
9. [Cron Jobs](#cron-jobs)
10. [Teams & Planner Provisioning](#teams--planner-provisioning)
11. [Telegram Setup](#telegram-setup)
12. [LLM Gateway Setup](#llm-gateway-setup)
13. [File Reference](#file-reference)
14. [Troubleshooting](#troubleshooting)

---

## Prerequisites

- [Node.js](https://nodejs.org/) v18 or later (for native `fetch` support)
- [OpenClaw](https://openclaw.dev) installed and running
- A Microsoft account (personal or work/school)
- (Optional) A Telegram bot token for Telegram delivery

---

## Installation

This skill lives at `~/.openclaw/workspace/skills/outlook-email/`. Clone or
copy this directory there, then install dependencies:

```bash
cd ~/.openclaw/workspace/skills/outlook-email
npm install
```

This installs:
- `@azure/msal-node` — Microsoft Authentication Library for token management
- `@modelcontextprotocol/sdk` — MCP server SDK

> **Note:** `node_modules`, token files, `outlook-summary.md`, and
> `run-state.json` are all `.gitignore`d and will not be committed.

---

## Azure App Registration

This is a one-time manual step per account type. You need one or two Azure app
registrations depending on your setup:

| Scenario | Apps needed |
|---|---|
| Personal email only | One personal app |
| Work email + Teams (same tenant) | One work/tenant app |
| Personal email + Teams | Two apps (personal + work) |
| Work email + Teams (same tenant, one app) | One work/tenant app — see [Single App for Both](#single-app-for-both-work-email--teams) |

---

### Personal Account App

Use this if your email is a personal Microsoft account (outlook.com, hotmail, live).

1. Go to [https://portal.azure.com](https://portal.azure.com) → **App registrations** → **New registration**
2. **Name:** `OpenClaw Assistant (Personal)`
3. **Supported account types:** `Personal Microsoft accounts only`
4. Click **Register**
5. Go to **Authentication** → **Add a platform** → **Mobile and desktop applications**
   - Tick: `https://login.microsoftonline.com/common/oauth2/nativeclient`
   - Click **Configure**
6. Under **Authentication**, set **Allow public client flows** → **Yes** → Save
7. Go to **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated permissions**
   Add:
   - `Mail.Read`, `Mail.ReadWrite`, `Mail.Send`
   - `Calendars.Read`, `Calendars.ReadWrite`
   - `offline_access`, `User.Read`
8. Go to **Overview** and copy the **Application (client) ID**
9. Paste it into `config.json` under `auth.personal.clientId`

---

### Work / School Account App (for Teams and Planner)

Use this for Teams, Planner, and work email. This app **must** be a single-tenant
registration (tied to your organization's Azure AD tenant).

1. Sign in to [https://portal.azure.com](https://portal.azure.com) with your **work account**
2. Go to **App registrations** → **New registration**
3. **Name:** `OpenClaw Assistant (Work)`
4. **Supported account types:** `Accounts in this organizational directory only (Single tenant)`
5. Click **Register**
6. Go to **Authentication** → **Add a platform** → **Mobile and desktop applications**
   - Tick: `https://login.microsoftonline.com/common/oauth2/nativeclient`
   - Click **Configure**
7. Under **Authentication**, set **Allow public client flows** → **Yes** → Save
8. Go to **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated permissions**
   Add:
   - `Mail.Read`, `Mail.ReadWrite`, `Mail.Send` *(if work email)*
   - `Calendars.Read`, `Calendars.ReadWrite` *(if work calendar)*
   - `ChannelMessage.Read.All`, `ChannelMessage.Send`
   - `Tasks.Read`, `Tasks.ReadWrite`
   - `offline_access`, `User.Read`
9. Click **Grant admin consent for [Your Org]** → Confirm
   > If you are not a tenant admin, ask your admin to grant consent.
10. Go to **Overview** and copy:
    - **Application (client) ID** → paste into `config.json` as `auth.work.clientId`
    - **Directory (tenant) ID** → paste into `config.json` as `auth.work.tenantId`

---

### Single App for Both Work Email + Teams

If your email is a work/school account (same tenant as Teams), you can use **one
single-tenant app** for everything — no personal app needed. In this case:

- Follow the [Work Account App](#work--school-account-app-for-teams-and-planner) steps above
- Include all Mail, Calendar, Channel, Tasks, and User permissions
- In `config.json`, only populate `auth.work` and leave `auth.personal` empty (or remove it)
- Authenticate once: `npm run auth:work`
- The MCP server and scripts will automatically use the work token for all operations

---

## Authentication

After filling in credentials in `config.json`, run the device code flow for each
account type you configured.

### Personal account
```bash
npm run auth:personal
# or: node scripts/auth.mjs --account=personal
```

### Work account
```bash
npm run auth:work
# or: node scripts/auth.mjs --account=work
```

If you only configured one account type, you can omit `--account` and it will
auto-detect.

For each run:
1. The script prints a URL and a short code
2. Open the URL in your browser and enter the code
3. Sign in with the appropriate Microsoft account
4. Return to the terminal — you'll see "Auth complete"

Token files written:
- Personal: `personal-tokens.json` + `personal-msal-cache.json`
- Work: `work-tokens.json` + `work-msal-cache.json`

**To re-authenticate** (if your token expires):
```bash
npm run auth:personal   # or auth:work
```

---

## MCP Server Registration

Register the MCP server with OpenClaw so the agent can use the tools
conversationally:

```bash
openclaw mcp set outlook '{
  "command": "node",
  "args": ["/home/node/.openclaw/workspace/skills/outlook-email/scripts/mcp-server.mjs"]
}'
```

> **Adjust the path** to match your actual install location.

Restart the OpenClaw agent to pick up the new MCP server:
```bash
openclaw restart
```

**Verify the server is registered:**
```bash
openclaw mcp list
```

---

## Configuration

Edit `config.json` to match your preferences:

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
        { "type": "telegram", "chatId": "YOUR_TELEGRAM_CHAT_ID" },
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
    "personal": {
      "clientId": "YOUR_PERSONAL_CLIENT_ID",
      "authority": "https://login.microsoftonline.com/consumers"
    },
    "work": {
      "clientId": "YOUR_WORK_CLIENT_ID",
      "tenantId": "YOUR_TENANT_ID"
    }
  }
}
```

If you only need one account type, you can leave the other block's values as
placeholders — the code will only use whichever account is actually authenticated.

### Key settings

| Setting | Description |
|---|---|
| `summary.wordCount` | Target word count for the summary prose |
| `summary.tone` | Writing tone instruction sent to the LLM |
| `summary.maxEmails` | Max emails fetched per run |
| `summary.unreadOnly` | `true` = only fetch unread emails |
| `summary.delivery.channels` | Array of delivery destinations (Telegram, Teams) |
| `todos.enabled` | `false` to disable to-do creation entirely |
| `todos.delivery.planId` | Microsoft Planner plan id (fill after Teams setup) |
| `todos.delivery.bucketId` | Planner bucket id (optional) |
| `auth.personal.clientId` | Azure app Client ID for personal Microsoft account |
| `auth.work.clientId` | Azure app Client ID for work/school account |
| `auth.work.tenantId` | Azure AD tenant/directory ID (from App Registration Overview) |
| `llm.modelOverride` | Override backend model for summaries; leave blank to use active model |

---

## Structured Workflows

### Email Summary

Run manually:
```bash
node scripts/run-summary.mjs
```

Dry run (preview output, skip delivery):
```bash
node scripts/run-summary.mjs --dry-run
```

The script:
1. Fetches emails using the personal account token (falls back to work if not configured)
2. Calls the LLM to produce a grouped summary
3. Delivers to all configured channels (Telegram uses email token; Teams uses work token)
4. Writes output to `outlook-summary.md` for follow-up
5. Updates `run-state.json` with `lastSummaryAt`

### To-Do Extraction

Run manually (after running the summary):
```bash
node scripts/run-todos.mjs
```

Dry run (preview tasks, don't create):
```bash
node scripts/run-todos.mjs --dry-run
```

The script:
1. Reads `outlook-summary.md`
2. Calls the LLM to extract action items as structured JSON
3. Creates each item as a Planner task (uses work account token)
4. Updates `run-state.json` with `lastTodosAt` and `createdTaskIds`

> **Telegram delivery** requires the `TELEGRAM_BOT_TOKEN` environment variable:
> ```bash
> export TELEGRAM_BOT_TOKEN="your_bot_token_here"
> node scripts/run-summary.mjs
> ```

---

## Skill Registration

Register `SKILL.md` as an OpenClaw skill so the agent uses the correct
instructions:

```bash
openclaw skill add outlook-email /home/node/.openclaw/workspace/skills/outlook-email/SKILL.md
```

> Adjust the path to your actual install location.

---

## Cron Jobs

Set up after everything is tested manually. The summary and to-do jobs run
every 6 hours, with the to-do job offset by 5 minutes.

### Summary (every 6 hours)

```bash
openclaw cron add \
  --name "Outlook Email Summary" \
  --every "6h" \
  --session isolated \
  --message "Run the email summary workflow by executing: node /home/node/.openclaw/workspace/skills/outlook-email/scripts/run-summary.mjs -- then confirm delivery completed or report any errors."
```

### To-Dos (every 6 hours, 5 min after summary)

```bash
openclaw cron add \
  --name "Email To-Dos" \
  --every "6h" \
  --delay "5m" \
  --session isolated \
  --message "Run the to-do creation workflow by executing: node /home/node/.openclaw/workspace/skills/outlook-email/scripts/run-todos.mjs -- then confirm tasks were created or report any errors."
```

**Verify cron jobs:**
```bash
openclaw cron list
```

> **Note:** Adjust `/home/node/...` to match your actual install path.

---

## Teams & Planner Provisioning

Run this after the MCP server is working, to discover your Teams IDs and fill
in `config.json`.

### Guided setup via the agent

1. **Start the conversation:**
   > "Set up Teams"

2. The agent will call `list_teams` and show your teams. Tell it which team to
   use for delivery.

3. The agent calls `list_channels` for that team. Tell it which channel to post
   to.

4. The agent writes `teamId` and `channelId` to `config.json`.

5. The agent calls `list_tasks` to find your Planner plans. Tell it which plan
   to use for to-dos.

6. The agent writes `planId` to `config.json`.

7. *(Optional)* Ask for buckets and tell the agent which one to use.

8. Preview the summary format:
   ```bash
   node scripts/run-summary.mjs --dry-run
   ```

### Manual setup (alternative)

Use the MCP tools via the OpenClaw agent:
- `list_teams` → note the `id` of your target team
- `list_channels teamId=<id>` → note the `id` of your target channel
- `list_tasks` → note the `planId` of your target plan

Then paste the IDs into `config.json`.

---

## Telegram Setup

1. Create a Telegram bot: message [@BotFather](https://t.me/BotFather) and use `/newbot`
2. Copy the bot token (format: `123456:ABC-DEF1234...`)
3. Get your chat ID: message [@userinfobot](https://t.me/userinfobot)
4. Set `summary.delivery.channels[].chatId` in `config.json`
5. Export the bot token before running summary scripts:
   ```bash
   export TELEGRAM_BOT_TOKEN="your_token_here"
   ```
   Or add it to your shell profile (`~/.bashrc`, `~/.zshrc`).

---

## LLM Gateway Setup

The summary and to-do scripts call OpenClaw's local OpenAI-compatible gateway.
Ensure the gateway is configured:

```bash
# Enable the chatCompletions endpoint
openclaw config patch --stdin <<'EOF'
{"gateway":{"http":{"endpoints":{"chatCompletions":{"enabled":true}}}}}
EOF

# Restart the gateway
openclaw gateway restart
```

The scripts read the gateway token directly from `~/.openclaw/openclaw.json`
(no environment variable needed). Ensure `gateway.auth.mode` is set to `"token"`.

To use a specific model for summaries, set `llm.modelOverride` in `config.json`:
```json
"llm": {
  "modelOverride": "openai/gpt-4o-mini"
}
```

---

## File Reference

| File | Description | Git-tracked |
|---|---|---|
| `SKILL.md` | Agent instruction layer | ✅ |
| `package.json` | Node.js package manifest | ✅ |
| `config.json` | User preferences and credentials | ✅ |
| `scripts/token.mjs` | Shared MSAL token helper | ✅ |
| `scripts/auth.mjs` | OAuth device code flow | ✅ |
| `scripts/mcp-server.mjs` | MCP stdio server | ✅ |
| `scripts/run-summary.mjs` | Email summary pipeline | ✅ |
| `scripts/run-todos.mjs` | To-do extraction pipeline | ✅ |
| `personal-tokens.json` | Personal account access token (auto-generated) | ❌ |
| `personal-msal-cache.json` | Personal account MSAL cache (auto-generated) | ❌ |
| `work-tokens.json` | Work account access token (auto-generated) | ❌ |
| `work-msal-cache.json` | Work account MSAL cache (auto-generated) | ❌ |
| `outlook-summary.md` | Latest summary output (auto-generated) | ❌ |
| `run-state.json` | Pipeline run state (auto-generated) | ❌ |
| `node_modules/` | npm dependencies | ❌ |

> Legacy `outlook-tokens.json` / `outlook-msal-cache.json` are still supported
> as fallback token files for personal accounts.

---

## Troubleshooting

### "No tokens found for personal account"
Run personal auth:
```bash
npm run auth:personal
```

### "No tokens found for work account"
Run work auth:
```bash
npm run auth:work
```

### "Access token for work account is expired"
Re-run auth:
```bash
npm run auth:work
```

### "auth.work.tenantId is required"
Your work app registration requires a tenant ID. Go to Azure Portal →
App registrations → your app → Overview → copy **Directory (tenant) ID**
and set it as `auth.work.tenantId` in `config.json`.

### "config.json has a placeholder clientId"
Edit `config.json` and replace `YOUR_PERSONAL_CLIENT_ID` or `YOUR_WORK_CLIENT_ID`
with your Azure app's actual Client ID.

### Graph API 403 Forbidden on Teams/Planner
Ensure admin consent was granted for `ChannelMessage.Read.All`, `ChannelMessage.Send`,
`Tasks.Read`, `Tasks.ReadWrite` in Azure Portal → App registrations → API permissions
→ Grant admin consent. These permissions require tenant admin approval.

### "openclaw.json not found"
The LLM scripts read the gateway token from `~/.openclaw/openclaw.json`.
Ensure OpenClaw is installed and `llm.gatewayConfigPath` in `config.json`
points to the right location.

### "TELEGRAM_BOT_TOKEN env var not set"
Export the variable before running:
```bash
export TELEGRAM_BOT_TOKEN="your_token"
```
Or add it to your shell profile.

### "Teams teamId or channelId not configured"
Run Teams provisioning or manually fill in `config.json`.

### MCP server not showing tools
1. Verify the path in `openclaw mcp set` is correct for your system
2. Run `openclaw mcp list` to confirm registration
3. Restart OpenClaw: `openclaw restart`
