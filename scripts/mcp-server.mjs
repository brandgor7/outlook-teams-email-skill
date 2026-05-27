/**
 * mcp-server.mjs — OpenClaw MCP stdio server for Outlook + Teams
 *
 * Provides MCP tools for email, calendar, and Teams interactions via
 * Microsoft Graph API. Registered with OpenClaw as the "outlook" server.
 *
 * Usage (registered in OpenClaw):
 *   openclaw mcp set outlook '{"command":"node","args":["<skill-root>/scripts/mcp-server.mjs"]}'
 */

import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
} from '@modelcontextprotocol/sdk/types.js';
import { getToken } from './token.mjs';

const GRAPH = 'https://graph.microsoft.com/v1.0';

async function graphRequest(token, path, options = {}) {
  const url = `${GRAPH}${path}`;
  const res = await fetch(url, {
    ...options,
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json',
      ...(options.headers ?? {}),
    },
    body: options.body ? JSON.stringify(options.body) : undefined,
  });

  if (!res.ok) {
    const errText = await res.text().catch(() => '');
    let errMsg = `Graph API error ${res.status}: ${res.statusText}`;
    try {
      const errJson = JSON.parse(errText);
      if (errJson?.error?.message) errMsg += ` — ${errJson.error.message}`;
    } catch (_) { /* ignore */ }
    throw new Error(errMsg);
  }

  const text = await res.text();
  return text ? JSON.parse(text) : null;
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

// ─── Tool definitions ────────────────────────────────────────────────────────

const TOOLS = [
  // Email tools
  {
    name: 'list_emails',
    description: 'List recent emails with subject, sender, preview, and id.',
    inputSchema: {
      type: 'object',
      properties: {
        top: { type: 'number', description: 'Max emails to return (default 20)', default: 20 },
        unread_only: { type: 'boolean', description: 'Only return unread emails', default: false },
        folder: { type: 'string', description: 'Mail folder name (default: inbox)', default: 'inbox' },
      },
    },
  },
  {
    name: 'get_email',
    description: 'Fetch the full body of one email by its id.',
    inputSchema: {
      type: 'object',
      properties: {
        id: { type: 'string', description: 'Email message id' },
      },
      required: ['id'],
    },
  },
  {
    name: 'send_email',
    description: 'Send an email.',
    inputSchema: {
      type: 'object',
      properties: {
        to: { type: 'string', description: 'Recipient email address' },
        subject: { type: 'string', description: 'Email subject' },
        body: { type: 'string', description: 'Email body (plain text or HTML)' },
        cc: { type: 'string', description: 'CC email address (optional)' },
      },
      required: ['to', 'subject', 'body'],
    },
  },
  {
    name: 'move_email',
    description: 'Move an email to a different folder.',
    inputSchema: {
      type: 'object',
      properties: {
        id: { type: 'string', description: 'Email message id' },
        folder: { type: 'string', description: 'Destination folder name (e.g. "deleteditems", "archive")' },
      },
      required: ['id', 'folder'],
    },
  },
  {
    name: 'mark_read',
    description: 'Mark an email as read or unread.',
    inputSchema: {
      type: 'object',
      properties: {
        id: { type: 'string', description: 'Email message id' },
        read: { type: 'boolean', description: 'true = mark read, false = mark unread' },
      },
      required: ['id', 'read'],
    },
  },
  // Calendar tools
  {
    name: 'list_calendar',
    description: 'List upcoming calendar events.',
    inputSchema: {
      type: 'object',
      properties: {
        days: { type: 'number', description: 'Number of days ahead to look (default 7)', default: 7 },
        top: { type: 'number', description: 'Max events to return (default 20)', default: 20 },
      },
    },
  },
  {
    name: 'create_event',
    description: 'Create a new calendar event.',
    inputSchema: {
      type: 'object',
      properties: {
        subject: { type: 'string', description: 'Event title' },
        start: { type: 'string', description: 'Start datetime (ISO 8601, e.g. 2026-06-01T10:00:00)' },
        end: { type: 'string', description: 'End datetime (ISO 8601)' },
        body: { type: 'string', description: 'Event description (optional)' },
        location: { type: 'string', description: 'Location (optional)' },
        attendees: {
          type: 'array',
          description: 'List of attendee email addresses (optional)',
          items: { type: 'string' },
        },
      },
      required: ['subject', 'start', 'end'],
    },
  },
  // Teams tools
  {
    name: 'list_teams',
    description: 'List Microsoft Teams the user belongs to.',
    inputSchema: { type: 'object', properties: {} },
  },
  {
    name: 'list_channels',
    description: 'List channels in a Teams team.',
    inputSchema: {
      type: 'object',
      properties: {
        teamId: { type: 'string', description: 'Team id (from list_teams)' },
      },
      required: ['teamId'],
    },
  },
  {
    name: 'post_teams_message',
    description: 'Post a message to a Teams channel.',
    inputSchema: {
      type: 'object',
      properties: {
        teamId: { type: 'string', description: 'Team id' },
        channelId: { type: 'string', description: 'Channel id' },
        message: { type: 'string', description: 'Message text (plain text or HTML)' },
      },
      required: ['teamId', 'channelId', 'message'],
    },
  },
  {
    name: 'list_tasks',
    description: 'List tasks from Microsoft Planner.',
    inputSchema: {
      type: 'object',
      properties: {
        planId: { type: 'string', description: 'Planner plan id (optional — lists all assigned tasks if omitted)' },
      },
    },
  },
  {
    name: 'create_task',
    description: 'Create a task in Microsoft Planner.',
    inputSchema: {
      type: 'object',
      properties: {
        title: { type: 'string', description: 'Task title' },
        planId: { type: 'string', description: 'Planner plan id' },
        bucketId: { type: 'string', description: 'Bucket id (optional)' },
        dueDate: { type: 'string', description: 'Due date (ISO 8601 date, optional)' },
        notes: { type: 'string', description: 'Task notes / description (optional)' },
      },
      required: ['title', 'planId'],
    },
  },
];

// ─── Tool handlers ────────────────────────────────────────────────────────────

async function handleListEmails({ top = 20, unread_only = false, folder = 'inbox' }) {
  const token = await getToken();
  const filter = unread_only ? '&$filter=isRead eq false' : '';
  const select = '$select=id,subject,from,receivedDateTime,isRead,bodyPreview';
  const data = await graphRequest(
    token,
    `/me/mailFolders/${encodeURIComponent(folder)}/messages?$top=${top}&${select}&$orderby=receivedDateTime desc${filter}`
  );
  return (data?.value ?? []).map((m) => ({
    id: m.id,
    subject: m.subject,
    from: m.from?.emailAddress?.address,
    fromName: m.from?.emailAddress?.name,
    receivedAt: m.receivedDateTime,
    isRead: m.isRead,
    preview: m.bodyPreview,
  }));
}

async function handleGetEmail({ id }) {
  const token = await getToken();
  const m = await graphRequest(token, `/me/messages/${encodeURIComponent(id)}`);
  const bodyContent = m.body?.contentType === 'html'
    ? stripHtml(m.body.content)
    : (m.body?.content ?? '');
  return {
    id: m.id,
    subject: m.subject,
    from: m.from?.emailAddress?.address,
    fromName: m.from?.emailAddress?.name,
    to: (m.toRecipients ?? []).map((r) => r.emailAddress?.address).join(', '),
    receivedAt: m.receivedDateTime,
    isRead: m.isRead,
    body: bodyContent,
  };
}

async function handleSendEmail({ to, subject, body, cc }) {
  const token = await getToken();
  const message = {
    subject,
    body: { contentType: 'Text', content: body },
    toRecipients: [{ emailAddress: { address: to } }],
  };
  if (cc) {
    message.ccRecipients = [{ emailAddress: { address: cc } }];
  }
  await graphRequest(token, '/me/sendMail', {
    method: 'POST',
    body: { message, saveToSentItems: true },
  });
  return { success: true, message: `Email sent to ${to}` };
}

async function handleMoveEmail({ id, folder }) {
  const token = await getToken();
  const result = await graphRequest(token, `/me/messages/${encodeURIComponent(id)}/move`, {
    method: 'POST',
    body: { destinationId: folder },
  });
  return { success: true, newId: result?.id };
}

async function handleMarkRead({ id, read }) {
  const token = await getToken();
  await graphRequest(token, `/me/messages/${encodeURIComponent(id)}`, {
    method: 'PATCH',
    body: { isRead: read },
  });
  return { success: true, isRead: read };
}

async function handleListCalendar({ days = 7, top = 20 }) {
  const token = await getToken();
  const now = new Date().toISOString();
  const end = new Date(Date.now() + days * 86400000).toISOString();
  const data = await graphRequest(
    token,
    `/me/calendarView?startDateTime=${now}&endDateTime=${end}&$top=${top}&$orderby=start/dateTime&$select=id,subject,start,end,location,organizer,attendees,bodyPreview`
  );
  return (data?.value ?? []).map((e) => ({
    id: e.id,
    subject: e.subject,
    start: e.start?.dateTime,
    end: e.end?.dateTime,
    location: e.location?.displayName,
    organizer: e.organizer?.emailAddress?.address,
    attendeeCount: (e.attendees ?? []).length,
    preview: e.bodyPreview,
  }));
}

async function handleCreateEvent({ subject, start, end, body, location, attendees }) {
  const token = await getToken();
  const event = {
    subject,
    start: { dateTime: start, timeZone: 'UTC' },
    end: { dateTime: end, timeZone: 'UTC' },
  };
  if (body) event.body = { contentType: 'Text', content: body };
  if (location) event.location = { displayName: location };
  if (attendees?.length) {
    event.attendees = attendees.map((email) => ({
      emailAddress: { address: email },
      type: 'required',
    }));
  }
  const result = await graphRequest(token, '/me/events', {
    method: 'POST',
    body: event,
  });
  return { success: true, id: result?.id, subject: result?.subject };
}

async function handleListTeams() {
  const token = await getToken();
  const data = await graphRequest(token, '/me/joinedTeams?$select=id,displayName,description');
  return (data?.value ?? []).map((t) => ({
    id: t.id,
    name: t.displayName,
    description: t.description,
  }));
}

async function handleListChannels({ teamId }) {
  const token = await getToken();
  const data = await graphRequest(
    token,
    `/teams/${encodeURIComponent(teamId)}/channels?$select=id,displayName,description`
  );
  return (data?.value ?? []).map((c) => ({
    id: c.id,
    name: c.displayName,
    description: c.description,
  }));
}

async function handlePostTeamsMessage({ teamId, channelId, message }) {
  const token = await getToken();
  const result = await graphRequest(
    token,
    `/teams/${encodeURIComponent(teamId)}/channels/${encodeURIComponent(channelId)}/messages`,
    {
      method: 'POST',
      body: { body: { contentType: 'html', content: message } },
    }
  );
  return { success: true, id: result?.id };
}

async function handleListTasks({ planId }) {
  const token = await getToken();
  let tasks;
  if (planId) {
    const data = await graphRequest(
      token,
      `/planner/plans/${encodeURIComponent(planId)}/tasks?$select=id,title,percentComplete,dueDateTime,priority,assignments`
    );
    tasks = data?.value ?? [];
  } else {
    // List tasks assigned to the current user across all plans
    const data = await graphRequest(
      token,
      '/me/planner/tasks?$select=id,title,percentComplete,dueDateTime,priority,planId'
    );
    tasks = data?.value ?? [];
  }
  return tasks.map((t) => ({
    id: t.id,
    title: t.title,
    percentComplete: t.percentComplete,
    dueDate: t.dueDateTime,
    priority: t.priority,
    planId: t.planId,
  }));
}

async function handleCreateTask({ title, planId, bucketId, dueDate, notes }) {
  const token = await getToken();

  // Get current user id for assignment
  const me = await graphRequest(token, '/me?$select=id');
  const userId = me?.id;

  const task = {
    title,
    planId,
    assignments: userId ? { [userId]: { '@odata.type': '#microsoft.graph.plannerAssignment', orderHint: ' !' } } : undefined,
  };
  if (bucketId) task.bucketId = bucketId;
  if (dueDate) task.dueDateTime = new Date(dueDate).toISOString();

  const created = await graphRequest(token, '/planner/tasks', {
    method: 'POST',
    body: task,
  });

  // Add notes via task details if provided
  if (notes && created?.id) {
    try {
      // Must use etag from the task detail to update
      const details = await graphRequest(token, `/planner/tasks/${created.id}/details`);
      const etag = details?.['@odata.etag'];
      if (etag) {
        await graphRequest(token, `/planner/tasks/${created.id}/details`, {
          method: 'PATCH',
          headers: { 'If-Match': etag },
          body: { description: notes },
        });
      }
    } catch (_) { /* notes are optional, don't fail task creation */ }
  }

  return { success: true, id: created?.id, title: created?.title };
}

// ─── Dispatch ─────────────────────────────────────────────────────────────────

const HANDLERS = {
  list_emails: handleListEmails,
  get_email: handleGetEmail,
  send_email: handleSendEmail,
  move_email: handleMoveEmail,
  mark_read: handleMarkRead,
  list_calendar: handleListCalendar,
  create_event: handleCreateEvent,
  list_teams: handleListTeams,
  list_channels: handleListChannels,
  post_teams_message: handlePostTeamsMessage,
  list_tasks: handleListTasks,
  create_task: handleCreateTask,
};

// ─── Server setup ─────────────────────────────────────────────────────────────

const server = new Server(
  { name: 'outlook', version: '1.0.0' },
  { capabilities: { tools: {} } }
);

server.setRequestHandler(ListToolsRequestSchema, async () => ({ tools: TOOLS }));

server.setRequestHandler(CallToolRequestSchema, async (request) => {
  const { name, arguments: args } = request.params;
  const handler = HANDLERS[name];

  if (!handler) {
    return {
      content: [{ type: 'text', text: `Unknown tool: ${name}` }],
      isError: true,
    };
  }

  try {
    const result = await handler(args ?? {});
    return {
      content: [{ type: 'text', text: JSON.stringify(result, null, 2) }],
    };
  } catch (err) {
    const isAuthError =
      err.message?.includes('expired') ||
      err.message?.includes('401') ||
      err.message?.includes('token');
    const msg = isAuthError
      ? `Auth error: ${err.message}\n\nRun: node scripts/auth.mjs`
      : `Error: ${err.message}`;
    return {
      content: [{ type: 'text', text: msg }],
      isError: true,
    };
  }
});

const transport = new StdioServerTransport();
await server.connect(transport);
