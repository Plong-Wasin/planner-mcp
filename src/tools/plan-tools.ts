import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { getGraphClient } from "../graph/client.js";
import { handleGraphError } from "../utils/error-handler.js";
import { log } from "../utils/logger.js";

export function registerPlanTools(server: McpServer): void {
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
}
