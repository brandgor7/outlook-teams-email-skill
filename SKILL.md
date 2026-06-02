---
name: outlook-teams-email-skill
description: >-
  Use when the user wants to read, search, draft, send, or manage Outlook email
  or calendar events; post to Microsoft Teams channels; or manage Planner tasks.
  Connects to the Microsoft Graph API via Azure app registration.
license: MIT
compatibility: >-
  Requires Node.js v18+, OpenClaw with MCP server support, a Microsoft Azure app
  registration, and outbound access to Microsoft Graph API
  (graph.microsoft.com). Telegram delivery requires the TELEGRAM_BOT_TOKEN
  environment variable.
metadata:
  author: brandgor7
  version: 2.0.0
  tags: outlook email teams planner calendar microsoft graph
allowed-tools: bash list_emails get_email send_email move_email mark_read list_calendar create_event list_teams list_channels post_teams_message list_tasks create_task
---

# Outlook + Teams Email Assistant

This skill connects OpenClaw to Microsoft Outlook and Teams via the Microsoft
Graph API. It provides two modes of operation:

---

## Conversational Mode (MCP tools)

Use MCP tools directly when the user asks ad-hoc questions or wants to act on
specific items. The MCP server (`mcp-server.mjs`) must be registered and running.

### Email
- **"Check my emails"** / **"What's in my inbox?"** → call `list_emails`
- **"Show me unread emails"** → call `list_emails` with `unread_only: true`
- **"Read that email from John"** / **"Open the email about the project"** → call `get_email` with the relevant id (look up id with `list_emails` first if needed)
- **"Draft a reply to..."** → call `get_email` first to read the original, then compose a reply and call `send_email`
- **"Send an email to..."** → call `send_email`
- **"Archive / move that email"** → call `move_email`
- **"Mark that as read"** → call `mark_read`

### Email Summary

After any `list_emails` or `get_email` call, always produce a formatted summary:

1. Read `references/email-summary.md` — this defines the output format and prompt structure.
2. Read `references/email-config.json` and substitute its values into the template:
   - `{{tone}}` → `prompt.tone`
   - `{{categories}}` → `prompt.categories`
   - `{{summary_max_words}}` → `prompt.summary_max_words`
   - `{{emails}}` → the full output returned by the MCP tool
3. Render and output the completed summary following the template exactly.

**If the user requests posting to Teams** (e.g. "post to Teams", "send to the channel", "share in Teams"):

4. Read `outputs.teams` from `references/email-config.json` to identify the target channel.
5. Call `list_teams` to find the team matching `outputs.teams.channel_name`, then call `list_channels` to resolve its `channelId` matching `outputs.teams.channel_id`.
6. Call `post_teams_message` with the formatted summary as the message body.

### Calendar
- **"What's on my calendar?"** / **"What do I have this week?"** → call `list_calendar`
- **"Create a meeting..."** / **"Schedule..."** → call `create_event`

### Teams
- **"Post to Teams..."** → call `post_teams_message`
- **"What teams am I in?"** → call `list_teams`
- **"What channels are in [team]?"** → call `list_channels`
- **"What tasks do I have?"** → call `list_tasks`
- **"Create a task..."** → call `create_task`

---

## Structured Mode (Scripts)

Use scripts for scheduled or bulk operations. Always suggest `--dry-run` first
so the user can preview before committing.

### Email Summary

The summary workflow is a two-step process — the script fetches emails and you
(the agent) produce the summary using the **Email Summary Format** below.

**Step 1 — Fetch emails** (script prints email data for you to read):
```bash
node skills/outlook-email/scripts/run-summary.mjs
```
Read the email data printed to stdout. Produce a summary following the
**Email Summary Format** section below. Write the result to `outlook-summary.md`
in the skill root.

**Step 2 — Deliver** (script posts `outlook-summary.md` to configured channels):
```bash
node skills/outlook-email/scripts/run-summary.mjs --deliver
```

**Preview delivery without sending:**
```bash
node skills/outlook-email/scripts/run-summary.mjs --deliver --dry-run
```

### To-Do Creation
- **"Create to-dos from my emails"** / scheduled trigger (after summary):
  ```bash
  node skills/outlook-email/scripts/run-todos.mjs
  ```
- **"Show me what tasks would be created"**:
  ```bash
  node skills/outlook-email/scripts/run-todos.mjs --dry-run
  ```

---

## Auth Failure Handling

- If any MCP tool returns an error containing "expired", "401", or "token" → tell
  the user to re-authenticate:
  ```bash
  node skills/outlook-email/scripts/auth.mjs
  ```
- Never silently swallow auth errors — always surface them to the user.
- After re-authentication, retry the operation automatically.

---

## Delivery

- Structured workflow output (summaries, task confirmations) is delivered to channels
  configured in `references/email-config.json` under `outputs`.
- Conversational output (direct answers, email reads, drafts) stays in the current
  chat session — do not post these to delivery channels unless explicitly requested.

---

## Provisioning Flows

### First-time Outlook Setup
1. User says "set up Outlook" → walk them through Azure app registration
   (see README.md for the exact steps)
2. User pastes their Client ID → write it to `config.json` under `auth.personal.clientId`
   (or `auth.work.clientId` for a work/school account)
3. Run auth:
   ```bash
   node skills/outlook-email/scripts/auth.mjs
   ```
4. Verify with `list_emails` — if it returns results, auth is working

### First-time Teams Setup
1. User says "set up Teams" → call `list_teams` and present the results
2. User selects a team → call `list_channels` for that team
3. User selects a channel → write `teamId` and `channelId` to `config.json`
4. Call `list_tasks` → present available Planner plans
5. User selects a plan → write `planId` to `config.json`
6. Optional: user selects a bucket → write `bucketId` to `config.json`
7. Fetch emails to confirm the output format:
   ```bash
   node skills/outlook-email/scripts/run-summary.mjs
   ```

---

## Notes

- **One token, all services**: Auth covers both Outlook and Teams since both
  use Microsoft Graph. Re-authenticating refreshes access to both.
- **Summarisation is agent-side**: Email summaries are produced by the agent
  (you) following the Email Summary Format above — no separate LLM API call.
  Only `run-todos.mjs` calls the OpenClaw gateway directly for task extraction.
- **Telegram delivery**: Requires `TELEGRAM_BOT_TOKEN` environment variable.
  Set it in your shell before running delivery.
