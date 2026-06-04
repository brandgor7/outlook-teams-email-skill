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
allowed-tools: bash list_emails get_email send_email move_email mark_read list_calendar create_event list_teams list_channels post_teams_message list_tasks create_task send_telegram
---

# Outlook + Teams Email Assistant

Connects to Microsoft Outlook and Teams via the Microsoft Graph API.

---

## Email

- **"Check my emails"** / **"What's in my inbox?"** → call `list_emails`
- **"Show me unread emails"** → call `list_emails` with `unread_only: true`
- **"Read that email from John"** → call `get_email` (look up id with `list_emails` first if needed)
- **"Draft a reply to..."** → call `get_email`, compose a reply, then call `send_email`
- **"Send an email to..."** → call `send_email`
- **"Archive / move that email"** → call `move_email`
- **"Mark that as read"** → call `mark_read`

After any `list_emails` or `get_email` call, always produce a formatted summary:

1. Read `references/email-summary.md` — defines the output format and prompt structure.
2. Read `references/email-config.json` and substitute its values into the template:
   - `{{tone}}` → `prompt.tone`
   - `{{categories}}` → `prompt.categories`
   - `{{summary_max_words}}` → `prompt.summary_max_words`
   - `{{emails}}` → the full output returned by the MCP tool
3. Produce the JSON summary. Present it to the user in a readable way in chat.

If the user requests posting to Teams (e.g. "post to Teams", "share in Teams"):

4. Read `outputs.teams` from `references/email-config.json` to identify the target channel.
5. Call `list_teams` to find the matching team, then `list_channels` to resolve `channelId`.
6. Call `post_teams_message` using the `teams_message` field from the JSON summary as the message body — **not** the raw JSON.

If the user requests posting to Telegram (e.g. "send to Telegram", "notify me"):

4. Read `outputs.telegram` from `references/email-config.json` to get `chat_id`.
5. Call `send_telegram` using the `teams_message` field from the JSON summary — **not** the raw JSON.

---

## Calendar

- **"What's on my calendar?"** / **"What do I have this week?"** → call `list_calendar`
- **"Create a meeting..."** / **"Schedule..."** → call `create_event`

---

## Teams

- **"Post to Teams..."** → call `post_teams_message`
- **"What teams am I in?"** → call `list_teams`
- **"What channels are in [team]?"** → call `list_channels`
- **"What tasks do I have?"** → call `list_tasks`
- **"Create a task..."** → call `create_task`

---

## Scheduled Workflows

When triggered on a schedule (e.g. cron), run the full summary + delivery workflow:

1. Call `list_emails` with `full_body: true` (and `unread_only`/`top` from `config.json`).
2. Produce a formatted summary following the Email summary steps above.
3. Deliver to configured channels using the `teams_message` field from the JSON summary:
   - Telegram: call `send_telegram` with the chat id from `references/email-config.json`
   - Teams: call `post_teams_message` with the team/channel ids from `references/email-config.json`
4. If `todos.enabled` is true in `config.json`: extract action items from the summary, then call `create_task` for each one using `todos.planId` and `todos.bucketId`.

---

## Auth Failure Handling

- If any MCP tool returns an error containing "expired", "401", or "token" → tell
  the user to re-authenticate:
  ```bash
  node scripts/auth.mjs --account=personal   # or --account=work
  ```
- Never silently swallow auth errors — always surface them to the user.
- After re-authentication, retry the operation automatically.

---

## Delivery

- Summary and task confirmation output is delivered to channels configured in
  `references/email-config.json` under `outputs`.
- Direct answers, email reads, and drafts stay in the current chat session —
  do not post these to delivery channels unless explicitly requested.

---

## Provisioning Flows

### First-time Outlook Setup
1. User says "set up Outlook" → walk them through Azure app registration
   (see README.md for the exact steps)
2. User pastes their Client ID → write it to `config.json` under `auth.personal.clientId`
   (or `auth.work.clientId` for a work/school account)
3. Run auth:
   ```bash
   node scripts/auth.mjs --account=personal
   ```
4. Verify with `list_emails` — if it returns results, auth is working

### First-time Teams Setup
1. User says "set up Teams" → call `list_teams` and present the results
2. User selects a team → call `list_channels` for that team
3. User selects a channel → write `teamId` and `channelId` to `config.json`
4. Call `list_tasks` → present available Planner plans
5. User selects a plan → write `planId` to `config.json`
6. Optional: user selects a bucket → write `bucketId` to `config.json`
7. Call `list_emails` with `full_body: true` to confirm email data looks correct.

---

## Notes

- **One token, all services**: Auth covers both Outlook and Teams since both
  use Microsoft Graph. Re-authenticating refreshes access to both.
- **Telegram delivery**: Requires `TELEGRAM_BOT_TOKEN` environment variable.
  Set it in your shell before running delivery.
