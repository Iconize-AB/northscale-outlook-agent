/**
 * Outlook Calendar MCP Server
 * Deployed on Vercel — Streamable HTTP (MCP spec) for clients like ElevenLabs.
 *
 * Required environment variables (set in Vercel dashboard):
 *   AZURE_TENANT_ID      – your Azure AD tenant ID
 *   AZURE_CLIENT_ID      – your Azure app client ID
 *   AZURE_CLIENT_SECRET  – your Azure app client secret
 *   OUTLOOK_USER_EMAIL   – the mailbox/calendar to manage (e.g. you@company.com)
 *   MCP_SECRET           – a random secret string to protect this endpoint
 */

import { ConfidentialClientApplication } from "@azure/msal-node";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { WebStandardStreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/webStandardStreamableHttp.js";
import * as z from "zod/v4";

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
  if (res.status === 204) return null;
  return res.json();
}

// ── Tool implementations ─────────────────────────────────────────────────────

async function checkAvailability({ date_from, date_to, slot_duration_minutes = 30 }) {
  const user = process.env.OUTLOOK_USER_EMAIL;
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

  const freeSlots = [];
  const slotMs = slot_duration_minutes * 60 * 1000;
  let cursor = new Date(date_from);
  const rangeEnd = new Date(date_to);

  while (cursor < rangeEnd) {
    const dayStart = new Date(cursor);
    dayStart.setHours(8, 0, 0, 0);
    const dayEnd = new Date(cursor);
    dayEnd.setHours(18, 0, 0, 0);

    let slotStart = new Date(Math.max(cursor.getTime(), dayStart.getTime()));

    while (slotStart.getTime() + slotMs <= dayEnd.getTime()) {
      const slotEnd = new Date(slotStart.getTime() + slotMs);
      const overlaps = busySlots.some((b) => slotStart < b.end && slotEnd > b.start);
      if (!overlaps) {
        freeSlots.push({ start: slotStart.toISOString(), end: slotEnd.toISOString() });
      }
      slotStart = new Date(slotStart.getTime() + slotMs);
    }

    cursor = new Date(cursor);
    cursor.setDate(cursor.getDate() + 1);
    cursor.setHours(0, 0, 0, 0);
  }

  return { free_slots: freeSlots.slice(0, 20), busy_count: busySlots.length };
}

async function createMeeting({
  subject,
  start,
  end,
  attendees = [],
  body_text = "",
  location = "",
  online_meeting = false,
}) {
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

async function cancelMeeting({ event_id }) {
  const user = process.env.OUTLOOK_USER_EMAIL;
  await graphRequest(`/users/${user}/events/${event_id}`, "DELETE");
  return { success: true, message: `Event ${event_id} has been cancelled.` };
}

function textResult(obj) {
  return {
    content: [{ type: "text", text: JSON.stringify(obj, null, 2) }],
  };
}

function createMcpServer() {
  const server = new McpServer({
    name: "outlook-calendar-mcp",
    version: "1.0.0",
  });

  server.registerTool(
    "check_availability",
    {
      description: "Check the Outlook calendar for free time slots between two dates.",
      inputSchema: {
        date_from: z.string().describe("Start of search range in ISO 8601, e.g. 2026-04-01T08:00:00Z"),
        date_to: z.string().describe("End of search range in ISO 8601, e.g. 2026-04-05T18:00:00Z"),
        slot_duration_minutes: z.number().optional().default(30).describe("Meeting length in minutes"),
      },
    },
    async (args) => textResult(await checkAvailability(args))
  );

  server.registerTool(
    "create_meeting",
    {
      description: "Create a calendar event in Outlook. Can invite attendees and create a Teams link.",
      inputSchema: {
        subject: z.string(),
        start: z.string().describe("Start in ISO 8601 UTC"),
        end: z.string().describe("End in ISO 8601 UTC"),
        attendees: z.array(z.string()).optional().default([]),
        body_text: z.string().optional().default(""),
        location: z.string().optional().default(""),
        online_meeting: z.boolean().optional().default(false),
      },
    },
    async (args) => textResult(await createMeeting(args))
  );

  server.registerTool(
    "cancel_meeting",
    {
      description: "Cancel (delete) an existing calendar event by its event ID.",
      inputSchema: {
        event_id: z.string().describe("The calendar event id from Microsoft Graph"),
      },
    },
    async (args) => textResult(await cancelMeeting(args))
  );

  return server;
}

/**
 * The MCP SDK rejects Streamable HTTP POSTs unless Accept lists BOTH
 * application/json and text/event-stream (spec). Many clients omit this and get 406,
 * which surfaces as ExceptionGroup in Python/ElevenLabs. Normalize before handling.
 */
function normalizeRequestForMcpTransport(request) {
  const headers = new Headers(request.headers);
  const method = request.method;

  if (method === "POST") {
    const accept = (headers.get("accept") || "").toLowerCase();
    if (!accept.includes("application/json") || !accept.includes("text/event-stream")) {
      headers.set("Accept", "application/json, text/event-stream");
      return new Request(request, { headers });
    }
  }

  if (method === "GET") {
    const accept = (headers.get("accept") || "").toLowerCase();
    if (!accept.includes("text/event-stream")) {
      headers.set("Accept", "text/event-stream");
      return new Request(request, { headers });
    }
  }

  return request;
}

/** Stateless transport — each invocation is isolated (fits Vercel serverless). */
function createTransport() {
  return new WebStandardStreamableHTTPServerTransport({
    sessionIdGenerator: undefined,
    // JSON responses avoid long-lived SSE streams (better on Vercel + picky HTTP clients).
    enableJsonResponse: true,
  });
}

function jsonUnauthorized() {
  return new Response(JSON.stringify({ error: "Unauthorized" }), {
    status: 401,
    headers: { "Content-Type": "application/json; charset=utf-8" },
  });
}

function corsResponse(response) {
  const headers = new Headers(response.headers);
  headers.set("Access-Control-Allow-Origin", "*");
  headers.set(
    "Access-Control-Expose-Headers",
    "mcp-session-id, mcp-protocol-version, mcp-trace-id"
  );
  return new Response(response.body, {
    status: response.status,
    statusText: response.statusText,
    headers,
  });
}

// ── Vercel serverless handler (Web Standard fetch + Streamable HTTP) ─────────

export default {
  async fetch(request) {
    if (request.method === "OPTIONS") {
      return new Response(null, {
        status: 204,
        headers: {
          "Access-Control-Allow-Origin": "*",
          "Access-Control-Allow-Methods": "GET, POST, DELETE, OPTIONS",
          "Access-Control-Allow-Headers":
            "Content-Type, Authorization, mcp-session-id, mcp-protocol-version, Last-Event-ID",
          "Access-Control-Max-Age": "86400",
        },
      });
    }

    const secret = process.env.MCP_SECRET;
    const authHeader = request.headers.get("authorization") || "";
    if (secret && authHeader !== `Bearer ${secret}`) {
      return jsonUnauthorized();
    }

    const transport = createTransport();
    const server = createMcpServer();

    try {
      await server.connect(transport);
      const req = normalizeRequestForMcpTransport(request);
      const response = await transport.handleRequest(req);
      return corsResponse(response);
    } catch (err) {
      console.error("MCP Streamable HTTP error:", err);
      return new Response(
        JSON.stringify({
          jsonrpc: "2.0",
          error: { code: -32603, message: err?.message || "Internal error" },
          id: null,
        }),
        {
          status: 500,
          headers: { "Content-Type": "application/json; charset=utf-8" },
        }
      );
    }
  },
};
