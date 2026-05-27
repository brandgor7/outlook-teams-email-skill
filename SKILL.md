# Outlook + Teams Email Assistant — Skill

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
- **"Run the summary"** / **"Summarise my emails"** / scheduled trigger:
  ```bash
  node skills/outlook-email/scripts/run-summary.mjs
  ```
- **"Preview the summary without sending"**:
  ```bash
  node skills/outlook-email/scripts/run-summary.mjs --dry-run
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

## Config Awareness

Before running structured workflows, check `config.json` for:
- `auth.clientId` — must not be `YOUR_CLIENT_ID_HERE`
- `summary.delivery.channels` — warn if Teams channel has empty `teamId`/`channelId`
- `todos.delivery.planId` — warn if empty (to-dos will fail without it)
- `llm.gatewayConfigPath` — must point to a valid `openclaw.json` with a gateway token

If any required values are missing, prompt the user to run the relevant setup step.

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

- Structured workflow output (summaries, task confirmations) is delivered to all
  channels configured in `config.json` under `summary.delivery.channels`.
- Conversational output (direct answers, email reads, drafts) stays in the current
  chat session — do not post these to delivery channels.

---

## Provisioning Flows

### First-time Outlook Setup
1. User says "set up Outlook" → walk them through Azure app registration
   (see README.md for the exact steps)
2. User pastes their Client ID → write it to `config.json` under `auth.clientId`
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
7. Run a dry-run summary to confirm everything looks right:
   ```bash
   node skills/outlook-email/scripts/run-summary.mjs --dry-run
   ```

---

## Notes

- **One token, all services**: Auth covers both Outlook and Teams since both
  use Microsoft Graph. Re-authenticating refreshes access to both.
- **LLM calls are local**: Summaries and task extraction call OpenClaw's local
  gateway endpoint — no external API key needed.
- **Telegram delivery**: Requires `TELEGRAM_BOT_TOKEN` environment variable.
  Set it in your shell before running summary scripts.
