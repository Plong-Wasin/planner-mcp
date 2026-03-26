import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { getGraphClient } from "../graph/client.js";
import { handleGraphError } from "../utils/error-handler.js";
import { log } from "../utils/logger.js";

export function registerUserTools(server: McpServer): void {
  // Register get me tool
  server.registerTool(
    "get_me",
    {
      description: "Get the current signed-in user's profile (id, displayName, mail, etc.)",
      inputSchema: {
        select: z
          .string()
          .optional()
          .describe("Comma-separated fields to return (e.g. 'id,displayName,mail,userPrincipalName'). Defaults to common fields."),
      },
    },
    async ({ select }) => {
      try {
        log("INFO", "get_me called", { select });
        const client = await getGraphClient();
        const endpoint = select ? `/me?$select=${encodeURIComponent(select)}` : "/me";
        const me = await client.api(endpoint).get();
        log("INFO", "Got current user", { id: me.id, displayName: me.displayName });
        return {
          content: [
            {
              type: "text",
              text: JSON.stringify(me, null, 2),
            },
          ],
        };
      } catch (error: any) {
        log("ERROR", "Error in get_me", {
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
}
