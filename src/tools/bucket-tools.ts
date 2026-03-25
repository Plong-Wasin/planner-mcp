import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { getGraphClient } from "../graph/client.js";
import { handleGraphError } from "../utils/error-handler.js";

export function registerBucketTools(server: McpServer): void {
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
}
