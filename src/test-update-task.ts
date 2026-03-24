import { Client } from "@microsoft/microsoft-graph-client";
import "isomorphic-fetch";
import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";

// Get the current file's directory
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Path for storing the access token
const tokenFilePath = path.join(__dirname, "../.access-token.txt");

// Logging configuration
const ENABLE_LOGGING = true;
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

// Helper function to get Graph client
function getGraphClient(): Client {
  // Try to read the stored access token
  let accessToken: string | null = null;
  
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

  if (!accessToken) {
    throw new Error("Not authenticated. Please provide access token via .access-token.txt file or GRAPH_ACCESS_TOKEN environment variable.");
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

// Main update task function
async function updateTask(params: {
  taskId: string;
  title?: string;
  description?: string;
  bucketId?: string;
  assignments?: string;
  dueDateTime?: string;
}) {
  const { taskId, title, description, bucketId, assignments, dueDateTime } = params;

  try {
    log("INFO", "update_task called", { taskId, title, description, bucketId, assignments, dueDateTime });

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
        throw new Error("Invalid JSON in assignments parameter");
      }
    }

    let task: any = null;

    // Only patch task if there are task-level updates
    if (Object.keys(taskData).length > 0) {
      log("INFO", "Updating task with data", { taskId, taskData });
      // Use If-Match header with the task's ETag
      task = await client
        .api(`/planner/tasks/${taskId}`)
        .header("If-Match", etag)
        .patch(taskData);
      log("INFO", "Task updated successfully", { taskId, task });
    } else {
      log("INFO", "No task-level updates, skipping task patch", { taskId });
      task = existingTask;
    }

    // If description or dueDateTime is provided, we need to update task details
    if (description || dueDateTime) {
      // First, get existing details to obtain ETag
      log("INFO", "Fetching existing task details", { taskId });
      const existingDetails = await client
        .api(`/planner/tasks/${taskId}/details`)
        .get();
      const etag = existingDetails["@odata.etag"];
      log("INFO", "Got existing details", { taskId, etag, existingDetails });

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
        .header("If-Match", etag)
        .patch(detailsData);
      log("INFO", "Task details updated successfully", { taskId });
    }

    return task;
  } catch (error) {
    log("ERROR", "Error in update_task", { taskId, error: (error as any)?.message || String(error), stack: (error as any)?.stack });
    throw new Error(handleGraphError(error));
  }
}

// CLI interface
async function main() {
  const args = process.argv.slice(2);
  
  if (args.length === 0) {
    console.log("Usage: node test-update-task.ts <taskId> [options]");
    console.log("");
    console.log("Options:");
    console.log("  --title <title>              Task title");
    console.log("  --description <description>  Task description");
    console.log("  --bucketId <bucketId>        Bucket ID");
    console.log("  --assignments <json>         Assignments as JSON string");
    console.log("  --dueDateTime <dateTime>    Due date/time (ISO 8601 format)");
    console.log("");
    console.log("Example:");
    console.log('  node test-update-task.ts PlAzDr6l8EGFoo-9yhCgL8kAFyHr --description "test"');
    process.exit(1);
  }

  const params: any = {};
  let taskId: string | undefined;

  for (let i = 0; i < args.length; i++) {
    const arg = args[i];
    
    if (arg === "--title" && args[i + 1]) {
      params.title = args[++i];
    } else if (arg === "--description" && args[i + 1]) {
      params.description = args[++i];
    } else if (arg === "--bucketId" && args[i + 1]) {
      params.bucketId = args[++i];
    } else if (arg === "--assignments" && args[i + 1]) {
      params.assignments = args[++i];
    } else if (arg === "--dueDateTime" && args[i + 1]) {
      params.dueDateTime = args[++i];
    } else if (!arg.startsWith("--") && !taskId) {
      taskId = arg;
    }
  }

  if (!taskId) {
    console.error("Error: taskId is required");
    process.exit(1);
  }

  params.taskId = taskId;

  try {
    console.log(`Updating task ${taskId}...`);
    const result = await updateTask(params);
    console.log("\n✅ Task updated successfully!");
    console.log("\nResult:");
    console.log(JSON.stringify(result, null, 2));
  } catch (error) {
    console.error("\n❌ Error updating task:");
    console.error(error instanceof Error ? error.message : String(error));
    process.exit(1);
  }
}

main();
