import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import { ClientSecretCredential } from "@azure/identity";
import { Client } from "@microsoft/microsoft-graph-client";
import "isomorphic-fetch";
import { fileURLToPath } from "url";
import path from "path";
import fs from "fs";

// Get the current file's directory
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Path for storing the access token
const tokenFilePath = path.join(__dirname, "../.access-token.txt");

// Logging configuration
const ENABLE_LOGGING = process.env.ENABLE_LOGGING === "true" || false;
const LOG_FILE = path.join(__dirname, "../planner-mcp.log");

// Helper function to log to both stderr and file
function log(level: string, message: string, data?: any) {
  if (!ENABLE_LOGGING) return;

  const timestamp = new Date().toISOString();
  const logEntry = {
    timestamp,
    level,
    message,
    ...(data !== undefined && { data }),
  };

  const logLine = JSON.stringify(logEntry);

  // Log to stderr (doesn't interfere with stdout communication)
  console.error(logLine);

  // Also append to log file
  try {
    fs.appendFileSync(LOG_FILE, logLine + "\n");
  } catch (err) {
    // Ignore file write errors
  }
}

// Type definitions for OAuth responses
interface DeviceCodeResponse {
  device_code: string;
  user_code: string;
  verification_uri: string;
  expires_in: number;
  interval: number;
  message?: string;
}

interface TokenResponse {
  access_token?: string;
  refresh_token?: string;
  error?: string;
  error_description?: string;
}

// Microsoft Graph API scopes for Planner
// Note: Planner requires Group.Read.All for listing plans, Tasks.ReadWrite for tasks,
// GroupMember.Read.All for listing group members, Directory.Read.All for reading user details,
// and offline_access for refresh tokens
const SCOPES = ["Group.Read.All", "Tasks.ReadWrite", "GroupMember.Read.All", "Directory.Read.All", "offline_access"];

// Default Client ID - Microsoft Graph Explorer
const DEFAULT_CLIENT_ID = "14d82eec-204b-4c2f-b7e8-296a70dab67e";

// Token storage (in production, this should be encrypted and persisted)
let accessToken: string | null = null;
let refreshToken: string | null = null;
let tokenExpiresAt: number | null = null; // Token expiration timestamp

// Store device code info for polling
let deviceCodeInfo: {
  deviceCode: string;
  clientId: string;
  tenantId: string;
  interval: number;
  expiresAt: number;
} | null = null;

// Try to read the stored access token and refresh token
try {
  if (fs.existsSync(tokenFilePath)) {
    const tokenData = fs.readFileSync(tokenFilePath, "utf8");
    try {
      // Try to parse as JSON first (new format)
      const parsedToken = JSON.parse(tokenData);
      accessToken = parsedToken.accessToken || parsedToken.token; // Support both old and new format
      refreshToken = parsedToken.refreshToken || null;
      tokenExpiresAt = parsedToken.expiresAt || null;
    } catch (parseError) {
      // Fall back to using the raw token (old format)
      accessToken = tokenData;
    }
  }
} catch (error) {
  console.error("Error reading access token file:", error);
}

// Alternatively, check if token is in environment variables
if (!accessToken && process.env.GRAPH_ACCESS_TOKEN) {
  accessToken = process.env.GRAPH_ACCESS_TOKEN;
}

// Create server instance
const server = new McpServer({
  name: "planner-mcp",
  version: "1.0.0",
});

// Helper function to refresh access token using refresh token
async function refreshAccessToken(): Promise<boolean> {
  if (!refreshToken) {
    log("INFO", "No refresh token available");
    return false;
  }

  try {
    log("INFO", "Attempting to refresh access token");

    const tokenResponse = await fetch(
      `https://login.microsoftonline.com/common/oauth2/v2.0/token`,
      {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: new URLSearchParams({
          client_id: DEFAULT_CLIENT_ID,
          grant_type: "refresh_token",
          refresh_token: refreshToken,
        }),
      },
    );

    const tokenData = (await tokenResponse.json()) as TokenResponse;

    if (tokenData.access_token) {
      accessToken = tokenData.access_token;
      refreshToken = tokenData.refresh_token || refreshToken; // Update refresh token if a new one is provided
      tokenExpiresAt = Date.now() + (60 * 60 * 1000); // Set expiration to 1 hour from now (default)

      // Save updated tokens to file
      try {
        fs.writeFileSync(
          tokenFilePath,
          JSON.stringify({
            accessToken: accessToken,
            refreshToken: refreshToken,
            expiresAt: tokenExpiresAt,
          }),
        );
      } catch (saveError) {
        console.error("Warning: Could not save refreshed token to file:", saveError);
      }

      log("INFO", "Access token refreshed successfully");
      return true;
    }

    if (tokenData.error === "invalid_grant" || tokenData.error === "refresh_token_expired") {
      log("ERROR", "Refresh token expired or invalid", { error: tokenData.error });
      // Clear tokens
      accessToken = null;
      refreshToken = null;
      tokenExpiresAt = null;
      try {
        fs.unlinkSync(tokenFilePath);
      } catch (e) {
        // Ignore error
      }
      return false;
    }

    log("ERROR", "Failed to refresh token", { error: tokenData.error });
    return false;
  } catch (error) {
    log("ERROR", "Error refreshing token", { error: (error as Error).message });
    return false;
  }
}

// Helper function to check if token is expired or about to expire (within 5 minutes)
function isTokenExpired(): boolean {
  if (!tokenExpiresAt) return false; // If we don't have expiration info, assume it's valid
  return Date.now() >= (tokenExpiresAt - 5 * 60 * 1000); // Expire 5 minutes early to be safe
}

// Helper function to get Graph client
async function getGraphClient(): Promise<Client> {
  // If no token in memory, try loading from file
  if (!accessToken) {
    try {
      if (fs.existsSync(tokenFilePath)) {
        const tokenData = fs.readFileSync(tokenFilePath, "utf8");
        try {
          const parsedToken = JSON.parse(tokenData);
          accessToken = parsedToken.accessToken || parsedToken.token;
          refreshToken = parsedToken.refreshToken || null;
          tokenExpiresAt = parsedToken.expiresAt || null;
          log("INFO", "Loaded token from file");
        } catch (parseError) {
          accessToken = tokenData;
          log("INFO", "Loaded token from file (legacy format)");
        }
      }
    } catch (error) {
      log("ERROR", "Failed to load token from file", { error: (error as Error).message });
    }
  }

  // Still no token after trying to load from file
  if (!accessToken) {
    throw new Error("Not authenticated. Please run authentication first or set GRAPH_ACCESS_TOKEN environment variable.");
  }

  if (isTokenExpired()) {
    log("INFO", "Access token expired, attempting to refresh");
    const refreshed = await refreshAccessToken();
    if (!refreshed || !accessToken) {
      throw new Error("Access token expired and refresh failed. Please re-authenticate.");
    }
  }

  const client = Client.init({
    authProvider: (done) => {
      done(null, accessToken!);
    },
  });

  return client;
}

// Helper function to handle Graph API errors
function handleGraphError(error: any): string {
  if (error?.code) {
    return `Graph API Error (${error.code}): ${error.message || "Unknown error"}`;
  }
  return `Error: ${error?.message || String(error)}`;
}

// Helper function to parse date string or relative date
function parseDate(dateInput: string): Date | null {
  if (!dateInput) return null;

  // Handle relative dates
  const now = new Date();
  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());

  switch (dateInput.toLowerCase()) {
    case "today":
      return today;
    case "tomorrow":
      return new Date(today.getTime() + 24 * 60 * 60 * 1000);
    case "yesterday":
      return new Date(today.getTime() - 24 * 60 * 60 * 1000);
    case "this-week": {
      const startOfWeek = new Date(today);
      startOfWeek.setDate(today.getDate() - today.getDay()); // Sunday
      return startOfWeek;
    }
    case "next-week": {
      const startOfNextWeek = new Date(today);
      startOfNextWeek.setDate(today.getDate() + (7 - today.getDay())); // Next Sunday
      return startOfNextWeek;
    }
    case "last-week": {
      const startOfLastWeek = new Date(today);
      startOfLastWeek.setDate(today.getDate() - today.getDay() - 7); // Last Sunday
      return startOfLastWeek;
    }
    default:
      // Try parsing as ISO 8601 or other date formats
      const parsed = new Date(dateInput);
      return isNaN(parsed.getTime()) ? null : parsed;
  }
}

// Helper function to check if a date is within range
function isDateInRange(dateString: string | null, startDate: Date, endDate: Date): boolean {
  if (!dateString) return false;
  const date = new Date(dateString);
  if (isNaN(date.getTime())) return false;
  return date >= startDate && date <= endDate;
}

// Register auth_start tool - initiates device code flow and shows the code/link
server.registerTool(
  "auth_start",
  {
    description:
      "Start the Planner authentication flow. This will display a verification URL and code that you need to use to sign in.",
    inputSchema: {
      clientId: z
        .string()
        .optional()
        .describe(
          "Azure AD application client ID (optional, uses Microsoft Graph Explorer by default)",
        ),
      tenantId: z
        .string()
        .optional()
        .describe("Azure AD tenant ID (optional, uses 'common' by default)"),
      force: z
        .boolean()
        .optional()
        .describe("Force re-authentication even if a valid token exists (default: false)"),
    },
  },
  async ({ clientId, tenantId, force }) => {
    try {
      // Check if we already have a valid token (unless force is true)
      if (!force) {
        if (accessToken && !isTokenExpired()) {
          return {
            content: [
              {
                type: "text",
                text: "✅ Already authenticated with a valid access token. No need to re-authenticate. Use force=true to re-authenticate.",
              },
            ],
          };
        }

        // Try to load from file if not in memory
        if (!accessToken) {
          try {
            if (fs.existsSync(tokenFilePath)) {
              const tokenData = fs.readFileSync(tokenFilePath, "utf8");
              try {
                const parsedToken = JSON.parse(tokenData);
                accessToken = parsedToken.accessToken || parsedToken.token;
                refreshToken = parsedToken.refreshToken || null;
                tokenExpiresAt = parsedToken.expiresAt || null;
              } catch (parseError) {
                accessToken = tokenData;
              }

              // Check if loaded token is valid
              if (accessToken && !isTokenExpired()) {
                return {
                  content: [
                    {
                      type: "text",
                      text: "✅ Already authenticated with a valid access token (loaded from storage). No need to re-authenticate. Use force=true to re-authenticate.",
                    },
                  ],
                };
              }
            }
          } catch (error) {
            log("ERROR", "Failed to check existing token", { error: (error as Error).message });
          }
        }
      }

      const client = clientId || DEFAULT_CLIENT_ID;
      const tenant = tenantId || "common";

      // Initiate device code flow
      const deviceCodeResponse = await fetch(
        `https://login.microsoftonline.com/${tenant}/oauth2/v2.0/devicecode`,
        {
          method: "POST",
          headers: { "Content-Type": "application/x-www-form-urlencoded" },
          body: new URLSearchParams({
            client_id: client,
            scope: SCOPES.join(" "),
          }),
        },
      );

      if (!deviceCodeResponse.ok) {
        throw new Error("Failed to initiate device code flow");
      }

      const deviceCode =
        (await deviceCodeResponse.json()) as DeviceCodeResponse;

      // Store device code info for polling
      deviceCodeInfo = {
        deviceCode: deviceCode.device_code,
        clientId: client,
        tenantId: tenant,
        interval: deviceCode.interval || 5,
        expiresAt: Date.now() + deviceCode.expires_in * 1000,
      };

      // Return instructions immediately without waiting
      return {
        content: [
          {
            type: "text",
            text: `Authentication Required!

1. Visit this URL: ${deviceCode.verification_uri}
2. Enter this code: ${deviceCode.user_code}
3. Sign in with your Microsoft account

After completing the authentication, use the "auth_poll" tool to check if authentication is complete.

⏱️ Code expires in ${Math.floor(deviceCode.expires_in / 60)} minutes`,
          },
        ],
      };
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Failed to start authentication: ${handleGraphError(error)}`,
          },
        ],
      };
    }
  },
);

// Register auth_poll tool - polls to check if authentication is complete
server.registerTool(
  "auth_poll",
  {
    description:
      "Check if the authentication is complete. Call this after using auth_start and completing the sign-in process on the Microsoft website.",
    inputSchema: {},
  },
  async () => {
    try {
      if (!deviceCodeInfo) {
        return {
          content: [
            {
              type: "text",
              text: "No authentication in progress. Please call auth_start first.",
            },
          ],
        };
      }

      // Check if device code has expired
      if (Date.now() > deviceCodeInfo.expiresAt) {
        deviceCodeInfo = null;
        return {
          content: [
            {
              type: "text",
              text: "❌ Authentication code has expired. Please call auth_start again.",
            },
          ],
        };
      }

      // Poll once for token
      const tokenResponse = await fetch(
        `https://login.microsoftonline.com/${deviceCodeInfo.tenantId}/oauth2/v2.0/token`,
        {
          method: "POST",
          headers: { "Content-Type": "application/x-www-form-urlencoded" },
          body: new URLSearchParams({
            client_id: deviceCodeInfo.clientId,
            grant_type: "urn:ietf:params:oauth:grant-type:device_code",
            device_code: deviceCodeInfo.deviceCode,
          }),
        },
      );

      const tokenData = (await tokenResponse.json()) as TokenResponse;

      if (tokenData.access_token) {
        accessToken = tokenData.access_token;
        refreshToken = tokenData.refresh_token || null;
        tokenExpiresAt = Date.now() + (60 * 60 * 1000); // Set expiration to 1 hour from now
        deviceCodeInfo = null; // Clear the device code info

        log("INFO", "Authentication successful", {
          hasRefreshToken: !!refreshToken,
          expiresAt: new Date(tokenExpiresAt!).toISOString(),
        });

        // Save tokens to file for future use
        try {
          fs.writeFileSync(
            tokenFilePath,
            JSON.stringify({
              accessToken: accessToken,
              refreshToken: refreshToken,
              expiresAt: tokenExpiresAt,
            }),
          );
          log("INFO", "Tokens saved to file successfully");
        } catch (saveError) {
          console.error("Warning: Could not save tokens to file:", saveError);
        }

        return {
          content: [
            {
              type: "text",
              text: "✅ Authentication successful! You can now use all Planner tools.",
            },
          ],
        };
      }

      if (tokenData.error === "authorization_pending") {
        const remainingSeconds = Math.floor(
          (deviceCodeInfo.expiresAt - Date.now()) / 1000,
        );
        const remainingMinutes = Math.floor(remainingSeconds / 60);
        return {
          content: [
            {
              type: "text",
              text: `⏳ Waiting for authentication... Please complete the sign-in on the Microsoft website and call auth_poll again.

Time remaining: ${remainingMinutes}m ${remainingSeconds % 60}s`,
            },
          ],
        };
      }

      if (tokenData.error === "authorization_declined") {
        deviceCodeInfo = null;
        return {
          content: [
            {
              type: "text",
              text: "❌ Authentication was declined. Please call auth_start to try again.",
            },
          ],
        };
      }

      if (tokenData.error === "expired_token") {
        deviceCodeInfo = null;
        return {
          content: [
            {
              type: "text",
              text: "❌ Authentication code has expired. Please call auth_start again.",
            },
          ],
        };
      }

      // Unknown error
      return {
        content: [
          {
            type: "text",
            text: `⚠️ Unexpected error: ${tokenData.error || "Unknown error"}. ${tokenData.error_description || ""}`,
          },
        ],
      };
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Authentication check failed: ${handleGraphError(error)}`,
          },
        ],
      };
    }
  },
);

// Register list plans tool
server.registerTool(
  "list_plans",
  {
    description: "List all Microsoft Planner plans for the current user",
    inputSchema: {},
  },
  async () => {
    try {
      const client = await getGraphClient();
      // Get all planner plans the user has access to
      const plans = await client.api("/me/planner/plans").get();

      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(plans, null, 2),
          },
        ],
      };
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: handleGraphError(error),
          },
        ],
      };
    }
  },
);

// Register get plan tool
server.registerTool(
  "get_plan",
  {
    description: "Get detailed information about a specific plan",
    inputSchema: {
      planId: z.string().describe("Plan ID (use list_plans to find ID)"),
    },
  },
  async ({ planId }) => {
    try {
      const client = await getGraphClient();
      const plan = await client.api(`/planner/plans/${planId}`).get();

      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(plan, null, 2),
          },
        ],
      };
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: handleGraphError(error),
          },
        ],
      };
    }
  },
);

// Register get plan details tool
server.registerTool(
  "get_plan_details",
  {
    description:
      "Get detailed information about a plan including category descriptions (labels)",
    inputSchema: {
      planId: z.string().describe("Plan ID (use list_plans to find ID)"),
    },
  },
  async ({ planId }) => {
    try {
      log("INFO", "get_plan_details called", { planId });

      const client = await getGraphClient();

      // Get plan details including category descriptions
      const planDetails = await client
        .api(`/planner/plans/${planId}/details`)
        .get();

      log("INFO", "Got plan details successfully", { planId, planDetails });

      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(planDetails, null, 2),
          },
        ],
      };
    } catch (error: any) {
      log("ERROR", "Error in get_plan_details", {
        planId,
        error: error?.message || String(error),
        stack: error?.stack,
      });
      return {
        content: [
          {
            type: "text",
            text: handleGraphError(error),
          },
        ],
      };
    }
  },
);

// Register list buckets tool
server.registerTool(
  "list_buckets",
  {
    description: "List all buckets in a plan",
    inputSchema: {
      planId: z.string().describe("Plan ID (use list_plans to find ID)"),
    },
  },
  async ({ planId }) => {
    try {
      const client = await getGraphClient();
      const buckets = await client
        .api(`/planner/plans/${planId}/buckets`)
        .get();

      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(buckets, null, 2),
          },
        ],
      };
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: handleGraphError(error),
          },
        ],
      };
    }
  },
);

// Register create bucket tool
server.registerTool(
  "create_bucket",
  {
    description: "Create a new bucket in a plan",
    inputSchema: {
      planId: z.string().describe("Plan ID"),
      name: z.string().describe("Bucket name"),
    },
  },
  async ({ planId, name }) => {
    try {
      const client = await getGraphClient();
      const bucket = await client.api("/planner/buckets").post({
        name: name,
        planId: planId,
      });

      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(bucket, null, 2),
          },
        ],
      };
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: handleGraphError(error),
          },
        ],
      };
    }
  },
);

// Register list tasks tool
server.registerTool(
  "list_tasks",
  {
    description: "List tasks in a plan or bucket with advanced filtering",
    inputSchema: {
      planId: z
        .string()
        .optional()
        .describe("Plan ID (optional - if not provided, lists tasks assigned to you)"),
      bucketId: z
        .string()
        .optional()
        .describe("Bucket ID (optional - filters by bucket within the plan)"),
      isArchived: z
        .boolean()
        .optional()
        .describe("Filter by archived status (true/false/undefined for both)"),
      percentComplete: z
        .number()
        .optional()
        .describe("Filter by percent complete (0-100), can use 0 for not started, 100 for completed"),
      priority: z
        .union([z.literal(1), z.literal(3), z.literal(5), z.literal(9)])
        .optional()
        .describe("Filter by priority (1=urgent, 3=important, 5=medium, 9=low)"),
      assignedToMe: z
        .boolean()
        .optional()
        .describe("Filter by tasks assigned to current user (only works with planId)"),
      // Date filters
      dueBefore: z
        .string()
        .optional()
        .describe("Due date before this date (ISO 8601: '2026-03-27' or relative: 'today', 'tomorrow', 'next-week')"),
      dueAfter: z
        .string()
        .optional()
        .describe("Due date after this date (ISO 8601: '2026-03-27' or relative: 'today', 'tomorrow', 'next-week')"),
      createdAfter: z
        .string()
        .optional()
        .describe("Created after this date (ISO 8601 or relative: 'today', 'this-week', 'last-week')"),
      createdBefore: z
        .string()
        .optional()
        .describe("Created before this date (ISO 8601 or relative: 'today', 'this-week', 'last-week')"),
      completedAfter: z
        .string()
        .optional()
        .describe("Completed after this date (ISO 8601 or relative)"),
      completedBefore: z
        .string()
        .optional()
        .describe("Completed before this date (ISO 8601 or relative)"),
      // Pagination and field selection
      limit: z
        .number()
        .optional()
        .describe("Maximum number of tasks to return (pagination)"),
      skip: z
        .number()
        .optional()
        .describe("Number of tasks to skip (for pagination, use with limit)"),
      fields: z
        .string()
        .optional()
        .describe("Comma-separated field names to return (e.g., 'id,title,bucketId,percentComplete') - reduces token usage"),
    },
  },
  async ({
    planId,
    bucketId,
    isArchived,
    percentComplete,
    priority,
    assignedToMe,
    dueBefore,
    dueAfter,
    createdAfter,
    createdBefore,
    completedAfter,
    completedBefore,
    limit,
    skip,
    fields,
  }) => {
    try {
      log("INFO", "list_tasks called", {
        planId,
        bucketId,
        isArchived,
        percentComplete,
        priority,
        assignedToMe,
        dueBefore,
        dueAfter,
        createdAfter,
        createdBefore,
        completedAfter,
        completedBefore,
        limit,
        skip,
        fields,
      });

      const client = await getGraphClient();
      let response: any;

      // Build query parameters (only $select is supported by Planner API)
      const queryParams: string[] = [];

      // Add $select for field selection (server-side) - Planner API supports this!
      if (fields) {
        queryParams.push(`$select=${encodeURIComponent(fields)}`);
      }

      // If planId is provided, use /planner/plans/{plan-id}/tasks to get ALL tasks in the plan
      if (planId) {
        let endpoint = `/planner/plans/${planId}/tasks`;
        if (queryParams.length > 0) {
          endpoint += `?${queryParams.join("&")}`;
        }
        response = await client.api(endpoint).get();
        log("INFO", "Fetched tasks from plan", { planId, count: response.value?.length });
      } else {
        // If no planId, list only tasks assigned to the current user
        let endpoint = "/me/planner/tasks";
        if (queryParams.length > 0) {
          endpoint += `?${queryParams.join("&")}`;
        }
        response = await client.api(endpoint).get();
        log("INFO", "Fetched tasks assigned to me", { count: response.value?.length });
      }

      // Client-side filtering
      if (response.value) {
        let filteredTasks = response.value;

        // Filter by bucketId
        if (bucketId) {
          filteredTasks = filteredTasks.filter((task: any) => task.bucketId === bucketId);
          log("INFO", "Filtered by bucketId", { bucketId, count: filteredTasks.length });
        }

        // Filter by isArchived
        if (isArchived !== undefined) {
          filteredTasks = filteredTasks.filter((task: any) => task.isArchived === isArchived);
          log("INFO", "Filtered by isArchived", { isArchived, count: filteredTasks.length });
        }

        // Filter by percentComplete
        if (percentComplete !== undefined) {
          filteredTasks = filteredTasks.filter((task: any) => task.percentComplete === percentComplete);
          log("INFO", "Filtered by percentComplete", { percentComplete, count: filteredTasks.length });
        }

        // Filter by priority
        if (priority !== undefined) {
          filteredTasks = filteredTasks.filter((task: any) => task.priority === priority);
          log("INFO", "Filtered by priority", { priority, count: filteredTasks.length });
        }

        // Filter by assignedToMe (only works if we have user context)
        if (assignedToMe !== undefined && assignedToMe) {
          // This requires knowing the current user's ID
          // For now, we'll filter tasks that have any assignments
          filteredTasks = filteredTasks.filter((task: any) => {
            const assignments = task.assignments || {};
            return Object.keys(assignments).length > 0;
          });
          log("INFO", "Filtered by assignedToMe", { count: filteredTasks.length });
        }

        // Filter by due date
        if (dueBefore) {
          const beforeDate = parseDate(dueBefore);
          if (beforeDate) {
            filteredTasks = filteredTasks.filter((task: any) => {
              if (!task.dueDateTime) return false;
              const dueDate = new Date(task.dueDateTime);
              return dueDate <= beforeDate;
            });
            log("INFO", "Filtered by dueBefore", { dueBefore, count: filteredTasks.length });
          }
        }

        if (dueAfter) {
          const afterDate = parseDate(dueAfter);
          if (afterDate) {
            filteredTasks = filteredTasks.filter((task: any) => {
              if (!task.dueDateTime) return false;
              const dueDate = new Date(task.dueDateTime);
              return dueDate >= afterDate;
            });
            log("INFO", "Filtered by dueAfter", { dueAfter, count: filteredTasks.length });
          }
        }

        // Filter by created date
        if (createdAfter) {
          const afterDate = parseDate(createdAfter);
          if (afterDate) {
            filteredTasks = filteredTasks.filter((task: any) => {
              if (!task.createdDateTime) return false;
              const createdDate = new Date(task.createdDateTime);
              return createdDate >= afterDate;
            });
            log("INFO", "Filtered by createdAfter", { createdAfter, count: filteredTasks.length });
          }
        }

        if (createdBefore) {
          const beforeDate = parseDate(createdBefore);
          if (beforeDate) {
            filteredTasks = filteredTasks.filter((task: any) => {
              if (!task.createdDateTime) return false;
              const createdDate = new Date(task.createdDateTime);
              return createdDate <= beforeDate;
            });
            log("INFO", "Filtered by createdBefore", { createdBefore, count: filteredTasks.length });
          }
        }

        // Filter by completed date
        if (completedAfter) {
          const afterDate = parseDate(completedAfter);
          if (afterDate) {
            filteredTasks = filteredTasks.filter((task: any) => {
              if (!task.completedDateTime) return false;
              const completedDate = new Date(task.completedDateTime);
              return completedDate >= afterDate;
            });
            log("INFO", "Filtered by completedAfter", { completedAfter, count: filteredTasks.length });
          }
        }

        if (completedBefore) {
          const beforeDate = parseDate(completedBefore);
          if (beforeDate) {
            filteredTasks = filteredTasks.filter((task: any) => {
              if (!task.completedDateTime) return false;
              const completedDate = new Date(task.completedDateTime);
              return completedDate <= beforeDate;
            });
            log("INFO", "Filtered by completedBefore", { completedBefore, count: filteredTasks.length });
          }
        }

        // Apply pagination (skip) - client-side since Planner API doesn't support $skip
        if (skip !== undefined && skip > 0) {
          filteredTasks = filteredTasks.slice(skip);
          log("INFO", "Applied skip", { skip, count: filteredTasks.length });
        }

        // Apply pagination (limit) - client-side since Planner API doesn't support $top
        if (limit !== undefined && limit > 0) {
          filteredTasks = filteredTasks.slice(0, limit);
          log("INFO", "Applied limit", { limit, count: filteredTasks.length });
        }

        // Note: $select (fields) is handled server-side via query parameters

        // Update response with filtered results
        response.value = filteredTasks;
        if ("@odata.count" in response) {
          response["@odata.count"] = filteredTasks.length;
        }

        log("INFO", "Final result", { count: filteredTasks.length });
      }

      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(response, null, 2),
          },
        ],
      };
    } catch (error: any) {
      log("ERROR", "Error in list_tasks", {
        error: error?.message || String(error),
        stack: error?.stack,
      });
      return {
        content: [
          {
            type: "text",
            text: handleGraphError(error),
          },
        ],
      };
    }
  },
);

// Register get task tool
server.registerTool(
  "get_task",
  {
    description: "Get detailed information about a specific task",
    inputSchema: {
      taskId: z.string().describe("Task ID"),
    },
  },
  async ({ taskId }) => {
    try {
      log("INFO", "get_task called", { taskId });
      const client = await getGraphClient();
      const task = await client.api(`/planner/tasks/${taskId}`).get();
      log("INFO", "Got task successfully", { taskId, task });

      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(task, null, 2),
          },
        ],
      };
    } catch (error: any) {
      log("ERROR", "Error in get_task", {
        taskId,
        error: error?.message || String(error),
        stack: error?.stack,
      });
      return {
        content: [
          {
            type: "text",
            text: handleGraphError(error),
          },
        ],
      };
    }
  },
);

// Register create task tool
server.registerTool(
  "create_task",
  {
    description: "Create a new task in a bucket",
    inputSchema: {
      planId: z.string().describe("Plan ID"),
      bucketId: z.string().describe("Bucket ID where to create the task"),
      title: z.string().describe("Task title"),
      description: z
        .string()
        .optional()
        .describe("Task description (optional)"),
      assignments: z
        .string()
        .optional()
        .describe("Task assignments as JSON string (optional)"),
      startDateTime: z
        .string()
        .optional()
        .describe("Start date in ISO 8601 format (optional)"),
      dueDateTime: z
        .string()
        .optional()
        .describe("Due date in ISO 8601 format (optional)"),
      percentComplete: z
        .number()
        .min(0)
        .max(100)
        .optional()
        .describe("Percentage of task completion 0-100 (optional)"),
      priority: z
        .union([z.literal(1), z.literal(3), z.literal(5), z.literal(9)])
        .optional()
        .describe("Priority: 1=urgent, 3=important, 5=medium, 9=low (optional)"),
      appliedCategories: z
        .string()
        .optional()
        .describe(
          "Labels/categories as JSON string with category names as keys and boolean values (e.g., '{\"category1\":true,\"category3\":true}') (optional)",
        ),
    },
  },
  async ({
    planId,
    bucketId,
    title,
    description,
    assignments,
    startDateTime,
    dueDateTime,
    percentComplete,
    priority,
    appliedCategories,
  }) => {
    try {
      log("INFO", "create_task called", {
        planId,
        bucketId,
        title,
        description,
        assignments,
        startDateTime,
        dueDateTime,
        percentComplete,
        priority,
        appliedCategories,
      });

      const client = await getGraphClient();

      const taskData: any = {
        planId: planId,
        bucketId: bucketId,
        title: title,
      };

      if (assignments) {
        try {
          taskData.assignments = JSON.parse(assignments);
        } catch (e) {
          return {
            content: [
              {
                type: "text",
                text: `Error: Invalid JSON in assignments parameter`,
              },
            ],
          };
        }
      }

      // Add startDateTime if provided (task property, NOT details)
      if (startDateTime !== undefined) {
        taskData.startDateTime = startDateTime;
      }

      // Add dueDateTime if provided (task property, NOT details)
      if (dueDateTime !== undefined) {
        taskData.dueDateTime = dueDateTime;
      }

      // Add percentComplete if provided
      if (percentComplete !== undefined) {
        taskData.percentComplete = percentComplete;
      }

      // Add priority if provided
      if (priority !== undefined) {
        taskData.priority = priority;
      }

      // Add appliedCategories if provided
      if (appliedCategories !== undefined) {
        try {
          taskData.appliedCategories = JSON.parse(appliedCategories);
        } catch (e) {
          return {
            content: [
              {
                type: "text",
                text: `Error: Invalid JSON in appliedCategories parameter`,
              },
            ],
          };
        }
      }

      log("INFO", "Creating task with data", { taskData });
      const task = await client.api("/planner/tasks").post(taskData);
      log("INFO", "Task created successfully", { task });

      // If description is provided, we need to update task details
      // Note: dueDateTime and startDateTime are task properties, NOT details properties
      if (description) {
        const taskId = task.id;
        log("INFO", "Creating task details", { taskId });

        const detailsData: any = {
          description: description,
        };

        // Create task details - this is a separate resource from the task itself
        // For new task details, we need to handle the ETag properly
        log("INFO", "Creating task details with data", { taskId, detailsData });

        try {
          // Try to get existing details first to check if they exist
          const existingDetails = await client
            .api(`/planner/tasks/${taskId}/details`)
            .get();

          // If details exist, update with If-Match header
          const detailsEtag = existingDetails["@odata.etag"];
          log("INFO", "Task details exist, updating with ETag", { taskId, detailsEtag });
          const details = await client
            .api(`/planner/tasks/${taskId}/details`)
            .header("If-Match", detailsEtag)
            .patch(detailsData);
          log("INFO", "Task details updated successfully", { taskId, details });
        } catch (getError: any) {
          // If details don't exist (404), create without If-Match header
          if (getError?.code === "NotFound" || getError?.statusCode === 404) {
            log("INFO", "Task details don't exist, creating new", { taskId });
            const details = await client
              .api(`/planner/tasks/${taskId}/details`)
              .patch(detailsData);
            log("INFO", "Task details created successfully", { taskId, details });
          } else {
            throw getError; // Re-throw other errors
          }
        }
      }

      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(task, null, 2),
          },
        ],
      };
    } catch (error: any) {
      log("ERROR", "Error in create_task", {
        error: error?.message || String(error),
        stack: error?.stack,
      });
      return {
        content: [
          {
            type: "text",
            text: handleGraphError(error),
          },
        ],
      };
    }
  },
);

// Register update task tool
server.registerTool(
  "update_task",
  {
    description: "Update an existing task",
    inputSchema: {
      taskId: z.string().describe("Task ID to update"),
      title: z.string().optional().describe("New task title (optional)"),
      description: z
        .string()
        .optional()
        .describe("New task description (optional)"),
      bucketId: z.string().optional().describe("New bucket ID (optional)"),
      assignments: z
        .string()
        .optional()
        .describe("New assignments as JSON string (optional)"),
      startDateTime: z
        .string()
        .optional()
        .describe("New start date in ISO 8601 format (optional)"),
      dueDateTime: z
        .string()
        .optional()
        .describe("New due date in ISO 8601 format (optional)"),
      percentComplete: z
        .number()
        .min(0)
        .max(100)
        .optional()
        .describe("Percentage of task completion 0-100 (optional)"),
      priority: z
        .union([z.literal(1), z.literal(3), z.literal(5), z.literal(9)])
        .optional()
        .describe("Priority: 1=urgent, 3=important, 5=medium, 9=low (optional)"),
      appliedCategories: z
        .string()
        .optional()
        .describe(
          "Labels/categories as JSON string with category names as keys and boolean values (e.g., '{\"category1\":true,\"category3\":true}') (optional)",
        ),
    },
  },
  async ({
    taskId,
    title,
    description,
    bucketId,
    assignments,
    startDateTime,
    dueDateTime,
    percentComplete,
    priority,
    appliedCategories,
  }) => {
    try {
      log("INFO", "update_task called", {
        taskId,
        title,
        description,
        bucketId,
        assignments,
        startDateTime,
        dueDateTime,
        percentComplete,
        priority,
        appliedCategories,
      });

      const client = await getGraphClient();

      // First, get the task to obtain its ETag
      log("INFO", "Fetching existing task", { taskId });
      const existingTask = await client.api(`/planner/tasks/${taskId}`).get();
      const etag = existingTask["@odata.etag"];
      log("INFO", "Got existing task", { taskId, etag, existingTask });

      const taskData: any = {};

      if (title !== undefined) {
        taskData.title = title;
      }

      if (bucketId !== undefined) {
        taskData.bucketId = bucketId;
      }

      if (assignments) {
        try {
          taskData.assignments = JSON.parse(assignments);
        } catch (e) {
          return {
            content: [
              {
                type: "text",
                text: `Error: Invalid JSON in assignments parameter`,
              },
            ],
          };
        }
      }

      // Add startDateTime if provided (task property, NOT details)
      if (startDateTime !== undefined) {
        taskData.startDateTime = startDateTime;
      }

      // Add dueDateTime if provided (task property, NOT details)
      if (dueDateTime !== undefined) {
        taskData.dueDateTime = dueDateTime;
      }

      // Add percentComplete if provided
      if (percentComplete !== undefined) {
        taskData.percentComplete = percentComplete;
      }

      // Add priority if provided
      if (priority !== undefined) {
        taskData.priority = priority;
      }

      // Add appliedCategories if provided
      if (appliedCategories !== undefined) {
        try {
          taskData.appliedCategories = JSON.parse(appliedCategories);
        } catch (e) {
          return {
            content: [
              {
                type: "text",
                text: `Error: Invalid JSON in appliedCategories parameter`,
              },
            ],
          };
        }
      }

      log("INFO", "Updating task with data", { taskId, taskData });
      // Use If-Match header with the task's ETag
      await client
        .api(`/planner/tasks/${taskId}`)
        .header("If-Match", etag)
        .patch(taskData);
      log("INFO", "Task updated successfully", { taskId });

      // If description is provided, we need to update task details
      // Note: dueDateTime and startDateTime are task properties, NOT details properties
      if (description) {
        // First, get existing details to obtain ETag
        log("INFO", "Fetching existing task details", { taskId });
        const existingDetails = await client
          .api(`/planner/tasks/${taskId}/details`)
          .get();
        const detailsEtag = existingDetails["@odata.etag"];
        log("INFO", "Got existing details", { taskId, detailsEtag, existingDetails });

        const detailsData: any = {
          description: description,
        };

        log("INFO", "Updating task details with data", { taskId, detailsData });
        // Use the ETag from existing details
        await client
          .api(`/planner/tasks/${taskId}/details`)
          .header("If-Match", detailsEtag)
          .patch(detailsData);
        log("INFO", "Task details updated successfully", { taskId });
      }

      // Fetch the updated task to return current data
      log("INFO", "Fetching updated task", { taskId });
      const updatedTask = await client.api(`/planner/tasks/${taskId}`).get();
      log("INFO", "Got updated task", { taskId, updatedTask });

      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(updatedTask, null, 2),
          },
        ],
      };
    } catch (error) {
      log("ERROR", "Error in update_task", {
        taskId,
        error: (error as any)?.message || String(error),
        stack: (error as any)?.stack,
      });
      return {
        content: [
          {
            type: "text",
            text: handleGraphError(error),
          },
        ],
      };
    }
  },
);

// Register delete task tool
server.registerTool(
  "delete_task",
  {
    description: "Delete a task",
    inputSchema: {
      taskId: z.string().describe("Task ID to delete"),
    },
  },
  async ({ taskId }) => {
    try {
      const client = await getGraphClient();

      // First, get the task to obtain its ETag
      const task = await client.api(`/planner/tasks/${taskId}`).get();
      const etag = task["@odata.etag"];

      // Delete with If-Match header using the task's ETag
      await client
        .api(`/planner/tasks/${taskId}`)
        .header("If-Match", etag)
        .delete();

      return {
        content: [
          {
            type: "text",
            text: `Task ${taskId} deleted successfully`,
          },
        ],
      };
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: handleGraphError(error),
          },
        ],
      };
    }
  },
);

// Register list group members tool
server.registerTool(
  "list_group_members",
  {
    description: "List all members of a Microsoft 365 group (users, contacts, devices, service principals, and other groups)",
    inputSchema: {
      groupId: z.string().describe("Group ID (use get_plan to find the group ID from plan's container.containerId)"),
      filter: z
        .string()
        .optional()
        .describe("OData filter expression (optional, e.g., \"displayName eq 'John Doe'\")"),
      search: z
        .string()
        .optional()
        .describe("Search string for displayName and description properties (optional)"),
      select: z
        .string()
        .optional()
        .describe("Comma-separated properties to return (optional, e.g., 'id,displayName,mail')"),
      top: z
        .number()
        .optional()
        .describe("Maximum number of members to return (optional, default: 100, max: 999)"),
    },
  },
  async ({ groupId, filter, search, select, top }) => {
    try {
      log("INFO", "list_group_members called", {
        groupId,
        filter,
        search,
        select,
        top,
      });

      const client = await getGraphClient();

      let endpoint = `/groups/${groupId}/members`;

      // Build query parameters
      const queryParams: string[] = [];

      if (filter) {
        queryParams.push(`$filter=${encodeURIComponent(filter)}`);
      }

      if (search) {
        queryParams.push(`$search=${encodeURIComponent(`"${search}"`)}`);
      }

      if (select) {
        queryParams.push(`$select=${encodeURIComponent(select)}`);
      }

      if (top) {
        queryParams.push(`$top=${top}`);
      }

      // Add query parameters to endpoint
      if (queryParams.length > 0) {
        endpoint += `?${queryParams.join("&")}`;
      }

      log("INFO", "Fetching group members", { endpoint });
      const members = await client.api(endpoint).get();
      log("INFO", "Fetched group members successfully", {
        count: members.value?.length,
      });

      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(members, null, 2),
          },
        ],
      };
    } catch (error: any) {
      log("ERROR", "Error in list_group_members", {
        groupId,
        error: error?.message || String(error),
        stack: error?.stack,
      });
      return {
        content: [
          {
            type: "text",
            text: handleGraphError(error),
          },
        ],
      };
    }
  },
);

// Main function to start the server
async function main() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error("Planner MCP Server running on stdio");
}

main().catch((error) => {
  console.error("Fatal error in main():", error);
  process.exit(1);
});
