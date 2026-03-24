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
const ENABLE_LOGGING = process.env.ENABLE_LOGGING === "true" || true;
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
// Note: Planner requires Group.Read.All for listing plans and Tasks.ReadWrite for tasks
const SCOPES = ["Group.Read.All", "Tasks.ReadWrite"];

// Default Client ID - Microsoft Graph Explorer
const DEFAULT_CLIENT_ID = "14d82eec-204b-4c2f-b7e8-296a70dab67e";

// Token storage (in production, this should be encrypted and persisted)
let accessToken: string | null = null;
let refreshToken: string | null = null;

// Store device code info for polling
let deviceCodeInfo: {
  deviceCode: string;
  clientId: string;
  tenantId: string;
  interval: number;
  expiresAt: number;
} | null = null;

// Try to read the stored access token
try {
  if (fs.existsSync(tokenFilePath)) {
    const tokenData = fs.readFileSync(tokenFilePath, "utf8");
    try {
      // Try to parse as JSON first (new format)
      const parsedToken = JSON.parse(tokenData);
      accessToken = parsedToken.token;
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

// Helper function to get Graph client
function getGraphClient(): Client {
  if (!accessToken) {
    throw new Error("Not authenticated. Please call auth_start first.");
  }

  const client = Client.init({
    authProvider: (done) => {
      done(null, accessToken);
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
    },
  },
  async ({ clientId, tenantId }) => {
    try {
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
        deviceCodeInfo = null; // Clear the device code info

        // Save token to file for future use
        try {
          fs.writeFileSync(
            tokenFilePath,
            JSON.stringify({ token: accessToken }),
          );
        } catch (saveError) {
          console.error("Warning: Could not save token to file:", saveError);
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
      const client = getGraphClient();
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
      const client = getGraphClient();
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
      const client = getGraphClient();
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
      orderHint: z
        .string()
        .optional()
        .describe("Order hint for positioning (optional)"),
    },
  },
  async ({ planId, name, orderHint }) => {
    try {
      const client = getGraphClient();
      const bucket = await client.api("/planner/buckets").post({
        name: name,
        planId: planId,
        orderHint: orderHint || " !",
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
    description: "List tasks in a plan or bucket",
    inputSchema: {
      planId: z
        .string()
        .optional()
        .describe("Plan ID (optional - if not provided, lists all tasks)"),
      bucketId: z
        .string()
        .optional()
        .describe("Bucket ID (optional - filters by bucket)"),
    },
  },
  async ({ planId, bucketId }) => {
    try {
      const client = getGraphClient();
      let endpoint = "/me/planner/tasks";

      if (bucketId) {
        const tasks = await client
          .api(endpoint)
          .filter(`bucketId eq '${bucketId}'`)
          .get();
        return {
          content: [
            {
              type: "text",
              text: JSON.stringify(tasks, null, 2),
            },
          ],
        };
      }

      if (planId) {
        const tasks = await client
          .api(endpoint)
          .filter(`planId eq '${planId}'`)
          .get();
        return {
          content: [
            {
              type: "text",
              text: JSON.stringify(tasks, null, 2),
            },
          ],
        };
      }

      // List all tasks
      const tasks = await client.api(endpoint).get();

      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(tasks, null, 2),
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
      const client = getGraphClient();
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
      dueDateTime: z
        .string()
        .optional()
        .describe("Due date in ISO 8601 format (optional)"),
    },
  },
  async ({
    planId,
    bucketId,
    title,
    description,
    assignments,
    dueDateTime,
  }) => {
    try {
      log("INFO", "create_task called", {
        planId,
        bucketId,
        title,
        description,
        assignments,
        dueDateTime,
      });

      const client = getGraphClient();

      const taskData: any = {
        planId: planId,
        bucketId: bucketId,
        title: title,
      };

      if (description) {
        taskData.description = description;
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

      if (dueDateTime) {
        // Due date needs to be set in task details
        taskData.dueDateTime = dueDateTime;
      }

      log("INFO", "Creating task with data", { taskData });
      const task = await client.api("/planner/tasks").post(taskData);
      log("INFO", "Task created successfully", { task });

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
      dueDateTime: z
        .string()
        .optional()
        .describe("New due date in ISO 8601 format (optional)"),
    },
  },
  async ({
    taskId,
    title,
    description,
    bucketId,
    assignments,
    dueDateTime,
  }) => {
    try {
      log("INFO", "update_task called", {
        taskId,
        title,
        description,
        bucketId,
        assignments,
        dueDateTime,
      });

      const client = getGraphClient();

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

      log("INFO", "Updating task with data", { taskId, taskData });
      // Use If-Match header with the task's ETag
      await client
        .api(`/planner/tasks/${taskId}`)
        .header("If-Match", etag)
        .patch(taskData);
      log("INFO", "Task updated successfully", { taskId });

      // If description or dueDateTime is provided, we need to update task details
      if (description || dueDateTime) {
        // First, get existing details to obtain ETag
        log("INFO", "Fetching existing task details", { taskId });
        const existingDetails = await client
          .api(`/planner/tasks/${taskId}/details`)
          .get();
        const detailsEtag = existingDetails["@odata.etag"];
        log("INFO", "Got existing details", { taskId, detailsEtag, existingDetails });

        const detailsData: any = {};

        if (description !== undefined) {
          detailsData.description = description;
        }

        if (dueDateTime !== undefined) {
          detailsData.dueDateTime = dueDateTime;
        }

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
      const client = getGraphClient();

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
