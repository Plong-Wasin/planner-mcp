import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { getGraphClient } from "../graph/client.js";
import { handleGraphError } from "../utils/error-handler.js";

export function registerGroupTools(server: McpServer): void {
  // Register list group members tool
  server.registerTool(
    "list_group_members",
    {
      description: "List all members of a Microsoft 365 group (users, contacts, devices, service principals, and other groups)",
      inputSchema: {
        groupId: z.string().describe("Group ID (use get_plan to find the group ID from plan's container.containerId)"),
        filter: z.string().optional().describe("OData filter expression (optional, e.g., \"displayName eq 'John Doe'\")"),
        search: z.string().optional().describe("Search string for displayName and description properties (optional)"),
        select: z.string().optional().describe("Comma-separated properties to return (optional, e.g., 'id,displayName,mail')"),
        top: z.number().optional().describe("Maximum number of members to return (optional, default: 100, max: 999)"),
      },
    },
    async ({ groupId, filter, search, select, top }) => {
      try {
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

        if (top !== undefined) {
          queryParams.push(`$top=${top}`);
        }

        if (queryParams.length > 0) {
          endpoint += `?${queryParams.join("&")}`;
        }

        const members = await client.api(endpoint).get();

        return {
          content: [
            {
              type: "text",
              text: JSON.stringify(members, null, 2),
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
}
