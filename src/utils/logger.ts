import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const ENABLE_LOGGING = process.env.ENABLE_LOGGING === "true" || false;
const LOG_FILE = path.join(__dirname, "../../planner-mcp.log");

export function log(level: string, message: string, data?: any) {
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
