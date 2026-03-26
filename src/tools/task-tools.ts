import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import Fuse from "fuse.js";
import { getGraphClient } from "../graph/client.js";
import { handleGraphError } from "../utils/error-handler.js";
import { log } from "../utils/logger.js";
import { parseDate } from "../utils/date-parser.js";

export function registerTaskTools(server: McpServer): void {
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
          .string()
          .optional()
          .describe("Filter by percent complete. Supports comparison operators: '>50', '>=50', '<100', '<=100', or exact value '50'. Use 0 for not started, 100 for completed"),
        priority: z
          .array(z.union([z.literal(1), z.literal(3), z.literal(5), z.literal(9)]))
          .optional()
          .describe("Filter by priority, supports multiple values (OR logic). e.g. [1,3] for urgent+important. Values: 1=urgent, 3=important, 5=medium, 9=low"),
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
        includeDetails: z
          .boolean()
          .optional()
          .describe("Include task details (description, checklist, references) via $expand=details"),
        assignedToUserIds: z
          .array(z.string())
          .optional()
          .describe("Filter tasks assigned to any of the specified user IDs (OR logic). e.g. ['user-id-1', 'user-id-2']"),
        appliedCategories: z
          .string()
          .optional()
          .describe("Filter tasks by categories, comma-separated (e.g., 'category1' or 'category1,category3'). Returns tasks that have ALL specified categories set to true."),
        search: z
          .string()
          .optional()
          .describe("Search tasks by keyword using fuzzy matching. Matches against title, and description if includeDetails is true."),
        searchThreshold: z
          .number()
          .min(0)
          .max(1)
          .optional()
          .describe("Fuzzy search sensitivity: 0.0 = exact match only, 1.0 = match anything. Default is 0.4."),
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
      includeDetails,
      assignedToUserIds,
      appliedCategories,
      search,
      searchThreshold,
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
          includeDetails,
          assignedToUserIds,
          appliedCategories,
          search,
          searchThreshold,
        });

        const client = await getGraphClient();
        let response: any;

        // Build query parameters (only $select is supported by Planner API)
        const queryParams: string[] = [];

        // Add $select for field selection (server-side) - Planner API supports this!
        // Auto-inject fields required by active client-side filters so they're always present.
        if (fields) {
          const selectedFields = new Set(fields.split(",").map((f) => f.trim()).filter(Boolean));
          if ((assignedToMe || (assignedToUserIds && assignedToUserIds.length > 0)) && !selectedFields.has("assignments")) {
            selectedFields.add("assignments");
          }
          if (bucketId && !selectedFields.has("bucketId")) {
            selectedFields.add("bucketId");
          }
          if (isArchived !== undefined && !selectedFields.has("isArchived")) {
            selectedFields.add("isArchived");
          }
          if (percentComplete !== undefined && !selectedFields.has("percentComplete")) {
            selectedFields.add("percentComplete");
          }
          if (priority !== undefined && !selectedFields.has("priority")) {
            selectedFields.add("priority");
          }
          if (appliedCategories && !selectedFields.has("appliedCategories")) {
            selectedFields.add("appliedCategories");
          }
          if ((dueBefore || dueAfter) && !selectedFields.has("dueDateTime")) {
            selectedFields.add("dueDateTime");
          }
          if ((createdAfter || createdBefore) && !selectedFields.has("createdDateTime")) {
            selectedFields.add("createdDateTime");
          }
          if ((completedAfter || completedBefore) && !selectedFields.has("completedDateTime")) {
            selectedFields.add("completedDateTime");
          }
          if (search && !selectedFields.has("title")) {
            selectedFields.add("title");
          }
          queryParams.push(`$select=${encodeURIComponent([...selectedFields].join(","))}`);
        }

        // Add $expand=details to fetch task details (description, checklist, references) inline
        if (includeDetails) {
          queryParams.push(`$expand=details`);
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

          // Filter by percentComplete (supports comparison operators: >50, >=50, <100, <=100, or exact value 50)
          if (percentComplete !== undefined) {
            const match = percentComplete.match(/^(>=|<=|>|<|=)?(\d+)$/);
            if (match) {
              const op = match[1] || "=";
              const val = parseInt(match[2], 10);
              filteredTasks = filteredTasks.filter((task: any) => {
                const pc = task.percentComplete ?? 0;
                if (op === ">") return pc > val;
                if (op === ">=") return pc >= val;
                if (op === "<") return pc < val;
                if (op === "<=") return pc <= val;
                return pc === val;
              });
            }
            log("INFO", "Filtered by percentComplete", { percentComplete, count: filteredTasks.length });
          }

          // Filter by priority (supports multiple values, OR logic)
          if (priority !== undefined && priority.length > 0) {
            const prioritySet = new Set(priority);
            filteredTasks = filteredTasks.filter((task: any) => prioritySet.has(task.priority));
            log("INFO", "Filtered by priority", { priority, count: filteredTasks.length });
          }

          // Filter by assignedToMe — fetch current user ID via /me
          if (assignedToMe !== undefined && assignedToMe) {
            const me = await client.api("/me").select("id").get();
            const myId = me.id as string;
            filteredTasks = filteredTasks.filter((task: any) => {
              const assignments = task.assignments || {};
              return myId in assignments;
            });
            log("INFO", "Filtered by assignedToMe", { myId, count: filteredTasks.length });
          }

          // Filter by multiple user IDs in assignments (OR logic — task assigned to ANY of them)
          if (assignedToUserIds && assignedToUserIds.length > 0) {
            filteredTasks = filteredTasks.filter((task: any) => {
              const assignments = task.assignments || {};
              return assignedToUserIds.some((uid: string) => uid in assignments);
            });
            log("INFO", "Filtered by assignedToUserIds", { assignedToUserIds, count: filteredTasks.length });
          }

          // Filter by appliedCategories - task must have ALL specified categories set to true
          if (appliedCategories) {
            const categoryList = appliedCategories.split(",").map((c) => c.trim()).filter(Boolean);
            filteredTasks = filteredTasks.filter((task: any) => {
              const categories = task.appliedCategories || {};
              return categoryList.every((cat) => categories[cat] === true);
            });
            log("INFO", "Filtered by appliedCategories", { appliedCategories, count: filteredTasks.length });
          }

          // Fuzzy search using Fuse.js (client-side)
          if (search) {
            const fuseKeys = includeDetails
              ? ["title", "details.description"]
              : ["title"];
            const fuse = new Fuse(filteredTasks, {
              keys: fuseKeys,
              threshold: searchThreshold ?? 0.4,
              includeScore: true,
              ignoreLocation: true,
            });
            filteredTasks = fuse.search(search).map((result) => result.item);
            log("INFO", "Fuzzy search applied", { search, threshold: searchThreshold ?? 0.4, keys: fuseKeys, count: filteredTasks.length });
          }

          // Filter by due date
          if (dueBefore) {
            const beforeDate = parseDate(dueBefore);
            if (beforeDate) {
              // If input has no time component (relative keyword or ISO date-only),
              // treat beforeDate as start-of-day and include the full day by using exclusive < next day
              const isDateOnly = !dueBefore.includes("T") && !dueBefore.includes(":");
              const exclusiveEnd = isDateOnly
                ? new Date(beforeDate.getTime() + 24 * 60 * 60 * 1000)
                : beforeDate;
              filteredTasks = filteredTasks.filter((task: any) => {
                if (!task.dueDateTime) return false;
                const dueDate = new Date(task.dueDateTime);
                return dueDate < exclusiveEnd;
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
        includeDetails: z
          .boolean()
          .optional()
          .describe("Include task details (description, checklist, references) via $expand=details"),
      },
    },
    async ({ taskId, includeDetails }) => {
      try {
        log("INFO", "get_task called", { taskId, includeDetails });
        const client = await getGraphClient();
        const endpoint = includeDetails
          ? `/planner/tasks/${taskId}?$expand=details`
          : `/planner/tasks/${taskId}`;
        const task = await client.api(endpoint).get();
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
            "Labels/categories as JSON string with category names as keys and boolean values. To ADD a tag set it to true (e.g., '{\"category1\":true}'). To REMOVE a tag you MUST set it to false (e.g., '{\"category1\":false}') — sending {} will NOT remove anything.",
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
            "Labels/categories as JSON string with category names as keys and boolean values. To ADD a tag set it to true (e.g., '{\"category1\":true}'). To REMOVE a tag you MUST set it to false (e.g., '{\"category1\":false}') — sending {} will NOT remove anything.",
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
        log("INFO", "delete_task called", { taskId });
        const client = await getGraphClient();

        // First, get the task to obtain its ETag
        const existingTask = await client.api(`/planner/tasks/${taskId}`).get();
        const etag = existingTask["@odata.etag"];

        // Use If-Match header with the task's ETag
        await client
          .api(`/planner/tasks/${taskId}`)
          .header("If-Match", etag)
          .delete();

        log("INFO", "Task deleted successfully", { taskId });

        return {
          content: [
            {
              type: "text",
              text: JSON.stringify({ success: true, taskId }, null, 2),
            },
          ],
        };
      } catch (error) {
        log("ERROR", "Error in delete_task", {
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
}
