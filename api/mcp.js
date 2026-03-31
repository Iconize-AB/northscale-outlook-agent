/**
 * Outlook Calendar MCP Server
 * Deployed on Vercel — exposes three tools to ElevenLabs agents:
 *   1. check_availability   — find free time slots
 *   2. create_meeting        — book a calendar event
 *   3. cancel_meeting        — delete a calendar event
 *
 * Required environment variables (set in Vercel dashboard):
 *   AZURE_TENANT_ID      – your Azure AD tenant ID
 *   AZURE_CLIENT_ID      – your Azure app client ID
 *   AZURE_CLIENT_SECRET  – your Azure app client secret
 *   OUTLOOK_USER_EMAIL   – the mailbox/calendar to manage (e.g. you@company.com)
 *   MCP_SECRET           – a random secret string to protect this endpoint
 */

import { ConfidentialClientApplication } = require("@azure/msal-node");

// ── Microsoft Graph helpers ──────────────────────────────────────────────────

const GRAPH_BASE = "https://graph.microsoft.com/v1.0";

async function getAccessToken() {
  const msalConfig = {
    auth: {
      clientId: process.env.AZURE_CLIENT_ID,
      authority: `https://login.microsoftonline.com/${process.env.AZURE_TENANT_ID}`,
      clientSecret: process.env.AZURE_CLIENT_SECRET,
    },
  };
  const cca = new ConfidentialClientApplication(msalConfig);
  const result = await cca.acquireTokenByClientCredential({
    scopes: ["https://graph.microsoft.com/.default"],
  });
  return result.accessToken;
}

async function graphRequest(path, method = "GET", body = null) {
  const token = await getAccessToken();
  const { default: fetch } = await import("node-fetch");
  const res = await fetch(`${GRAPH_BASE}${path}`, {
    method,
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
    body: body ? JSON.stringify(body) : undefined,
  });
  if (!res.ok) {
    const err = await res.text();
    throw new Error(`Graph API error ${res.status}: ${err}`);
  }
  if (res.status === 204) return null; // No content (e.g. DELETE)
  return res.json();
}

// ── Tool implementations ─────────────────────────────────────────────────────

/**
 * check_availability
 * Returns free time slots within a given date range.
 */
async function checkAvailability({ date_from, date_to, slot_duration_minutes = 30 }) {
  const user = process.env.OUTLOOK_USER_EMAIL;

  // Fetch existing events in the range
  const from = new Date(date_from).toISOString();
  const to = new Date(date_to).toISOString();

  const data = await graphRequest(
    `/users/${user}/calendarView?startDateTime=${from}&endDateTime=${to}&$select=subject,start,end&$orderby=start/dateTime&$top=50`
  );

  const events = data.value || [];
  const busySlots = events.map((e) => ({
    start: new Date(e.start.dateTime + (e.start.timeZone === "UTC" ? "Z" : "")),
    end: new Date(e.end.dateTime + (e.end.timeZone === "UTC" ? "Z" : "")),
  }));

  // Find free slots (simple gap analysis between 08:00–18:00 each day)
  const freeSlots = [];
  const slotMs = slot_duration_minutes * 60 * 1000;

  let cursor = new Date(date_from);
  const rangeEnd = new Date(date_to);

  while (cursor < rangeEnd) {
    // Only consider working hours: 08:00–18:00
    const dayStart = new Date(cursor);
    dayStart.setHours(8, 0, 0, 0);
    const dayEnd = new Date(cursor);
    dayEnd.setHours(18, 0, 0, 0);

    let slotStart = new Date(Math.max(cursor.getTime(), dayStart.getTime()));

    while (slotStart.getTime() + slotMs <= dayEnd.getTime()) {
      const slotEnd = new Date(slotStart.getTime() + slotMs);
      const overlaps = busySlots.some(
        (b) => slotStart < b.end && slotEnd > b.start
      );
      if (!overlaps) {
        freeSlots.push({
          start: slotStart.toISOString(),
          end: slotEnd.toISOString(),
        });
      }
      slotStart = new Date(slotStart.getTime() + slotMs);
    }

    // Advance to next day
    cursor = new Date(cursor);
    cursor.setDate(cursor.getDate() + 1);
    cursor.setHours(0, 0, 0, 0);
  }

  return {
    free_slots: freeSlots.slice(0, 20), // cap at 20 slots
    busy_count: busySlots.length,
  };
}

/**
 * create_meeting
 * Creates a calendar event and optionally invites attendees.
 */
async function createMeeting({ subject, start, end, attendees = [], body_text = "", location = "", online_meeting = false }) {
  const user = process.env.OUTLOOK_USER_EMAIL;

  const eventPayload = {
    subject,
    body: { contentType: "Text", content: body_text },
    start: { dateTime: start, timeZone: "UTC" },
    end: { dateTime: end, timeZone: "UTC" },
    location: location ? { displayName: location } : undefined,
    attendees: attendees.map((email) => ({
      emailAddress: { address: email },
      type: "required",
    })),
    isOnlineMeeting: online_meeting,
    onlineMeetingProvider: online_meeting ? "teamsForBusiness" : undefined,
  };

  const created = await graphRequest(`/users/${user}/events`, "POST", eventPayload);

  return {
    id: created.id,
    subject: created.subject,
    start: created.start.dateTime,
    end: created.end.dateTime,
    web_link: created.webLink,
    online_meeting_url: created.onlineMeeting?.joinUrl || null,
  };
}

/**
 * cancel_meeting
 * Deletes/cancels a calendar event by its ID.
 */
async function cancelMeeting({ event_id }) {
  const user = process.env.OUTLOOK_USER_EMAIL;
  await graphRequest(`/users/${user}/events/${event_id}`, "DELETE");
  return { success: true, message: `Event ${event_id} has been cancelled.` };
}

// ── MCP Protocol handler ─────────────────────────────────────────────────────

const TOOLS = [
  {
    name: "check_availability",
    description:
      "Check the Outlook calendar for free time slots between two dates. Returns a list of available time windows.",
    inputSchema: {
      type: "object",
      properties: {
        date_from: {
          type: "string",
          description: "Start of the search range in ISO 8601 format, e.g. '2026-04-01T08:00:00Z'",
        },
        date_to: {
          type: "string",
          description: "End of the search range in ISO 8601 format, e.g. '2026-04-05T18:00:00Z'",
        },
        slot_duration_minutes: {
          type: "number",
          description: "Desired meeting duration in minutes. Defaults to 30.",
        },
      },
      required: ["date_from", "date_to"],
    },
  },
  {
    name: "create_meeting",
    description:
      "Create a calendar event (meeting) in Outlook. Can invite attendees and create a Teams link.",
    inputSchema: {
      type: "object",
      properties: {
        subject: { type: "string", description: "Meeting title" },
        start: { type: "string", description: "Start time in ISO 8601 UTC, e.g. '2026-04-02T10:00:00Z'" },
        end: { type: "string", description: "End time in ISO 8601 UTC, e.g. '2026-04-02T10:30:00Z'" },
        attendees: {
          type: "array",
          items: { type: "string" },
          description: "List of attendee email addresses",
        },
        body_text: { type: "string", description: "Optional meeting description/agenda" },
        location: { type: "string", description: "Optional physical location" },
        online_meeting: {
          type: "boolean",
          description: "If true, creates a Microsoft Teams meeting link",
        },
      },
      required: ["subject", "start", "end"],
    },
  },
  {
    name: "cancel_meeting",
    description: "Cancel (delete) an existing calendar event by its event ID.",
    inputSchema: {
      type: "object",
      properties: {
        event_id: {
          type: "string",
          description: "The unique ID of the calendar event to cancel (returned by create_meeting or from the calendar)",
        },
      },
      required: ["event_id"],
    },
  },
];

// ── Vercel serverless handler ────────────────────────────────────────────────

module.exports = async function handler(req, res) {
  // Simple bearer-token guard
  const authHeader = req.headers["authorization"] || "";
  const secret = process.env.MCP_SECRET;
  if (secret && authHeader !== `Bearer ${secret}`) {
    return res.status(401).json({ error: "Unauthorized" });
  }

  if (req.method !== "POST") {
    return res.status(405).json({ error: "Method not allowed" });
  }

  let body;
  try {
    body = typeof req.body === "string" ? JSON.parse(req.body) : req.body;
  } catch {
    return res.status(400).json({ error: "Invalid JSON" });
  }

  const { jsonrpc, id, method, params } = body;

  // MCP JSON-RPC dispatch
  try {
    if (method === "initialize") {
      return res.json({
        jsonrpc: "2.0",
        id,
        result: {
          protocolVersion: "2024-11-05",
          capabilities: { tools: {} },
          serverInfo: { name: "outlook-calendar-mcp", version: "1.0.0" },
        },
      });
    }

    if (method === "tools/list") {
      return res.json({ jsonrpc: "2.0", id, result: { tools: TOOLS } });
    }

    if (method === "tools/call") {
      const { name, arguments: args } = params;
      let result;

      if (name === "check_availability") result = await checkAvailability(args);
      else if (name === "create_meeting") result = await createMeeting(args);
      else if (name === "cancel_meeting") result = await cancelMeeting(args);
      else throw new Error(`Unknown tool: ${name}`);

      return res.json({
        jsonrpc: "2.0",
        id,
        result: {
          content: [{ type: "text", text: JSON.stringify(result, null, 2) }],
        },
      });
    }

    // Fallback for unsupported methods
    return res.json({
      jsonrpc: "2.0",
      id,
      error: { code: -32601, message: "Method not found" },
    });
  } catch (err) {
    console.error("MCP handler error:", err);
    return res.status(500).json({
      jsonrpc: "2.0",
      id,
      error: { code: -32603, message: err.message },
    });
  }
};
