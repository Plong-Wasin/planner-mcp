import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { startDeviceCodeFlow, pollForToken } from "../auth/device-code-flow.js";

export function registerAuthTools(server: McpServer): void {
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
      const result = await startDeviceCodeFlow(clientId, tenantId, force);
      return {
        content: [
          {
            type: "text",
            text: result.message,
          },
        ],
      };
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
      const result = await pollForToken();
      return {
        content: [
          {
            type: "text",
            text: result.message,
          },
        ],
      };
    },
  );
}
