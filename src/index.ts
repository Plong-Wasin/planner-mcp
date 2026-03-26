import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";

// Import tool registrars
import { registerAuthTools } from "./tools/auth-tools.js";
import { registerPlanTools } from "./tools/plan-tools.js";
import { registerBucketTools } from "./tools/bucket-tools.js";
import { registerTaskTools } from "./tools/task-tools.js";
import { registerGroupTools } from "./tools/group-tools.js";
import { registerUserTools } from "./tools/user-tools.js";

// Create server instance
const server = new McpServer({
  name: "planner-mcp",
  version: "1.0.0",
});

// Register all tools
registerAuthTools(server);
registerPlanTools(server);
registerBucketTools(server);
registerTaskTools(server);
registerGroupTools(server);
registerUserTools(server);

// Start the server
async function main() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error("Planner MCP server running on stdio");
}

main().catch((error) => {
  console.error("Fatal error in main():", error);
  process.exit(1);
});
