# Test Update Task CLI Tool

This is a standalone command-line tool for testing the `update_task` functionality without going through MCP.

## Prerequisites

- You must have a valid access token stored in `.access-token.txt` file in the project root
- Or set the `GRAPH_ACCESS_TOKEN` environment variable

## Installation

The test script requires `tsx` to run TypeScript files directly. Install it if you haven't already:

```bash
npm install -g tsx
```

Or install it as a dev dependency:

```bash
npm install --save-dev tsx
```

## Usage

### Basic Syntax

```bash
npm run test-update-task -- <taskId> [options]
```

**Important:** Use `--` after the script name to separate npm flags from script arguments.

### Options

| Option | Description |
|--------|-------------|
| `--title <title>` | Task title |
| `--description <description>` | Task description |
| `--bucketId <bucketId>` | Bucket ID |
| `--assignments <json>` | Assignments as JSON string |
| `--dueDateTime <dateTime>` | Due date/time (ISO 8601 format) |

### Examples

#### Update task description only

```bash
npm run test-update-task -- PlAzDr6l8EGFoo-9yhCgL8kAFyHr --description "test"
```

#### Update task title and description

```bash
npm run test-update-task -- PlAzDr6l8EGFoo-9yhCgL8kAFyHr --title "New Title" --description "Updated description"
```

#### Update task with assignments

```bash
npm run test-update-task -- PlAzDr6l8EGFoo-9yhCgL8kAFyHr --assignments '{"user-id@domain.com":{"@odata.type":"#microsoft.graph.plannerAssignment","orderHint":" !"}}'
```

#### Update task with due date

```bash
npm run test-update-task -- PlAzDr6l8EGFoo-9yhCgL8kAFyHr --dueDateTime "2026-03-25T17:00:00Z"
```

#### Update all fields

```bash
npm run test-update-task -- PlAzDr6l8EGFoo-9yhCgL8kAFyHr \
  --title "Complete Task" \
  --description "This is a test task" \
  --bucketId "bucketId123" \
  --dueDateTime "2026-03-25T17:00:00Z"
```

## How It Works

1. The script reads your access token from `.access-token.txt` or the `GRAPH_ACCESS_TOKEN` environment variable
2. It fetches the existing task to get its ETag (required for optimistic concurrency)
3. It updates the task fields (title, bucketId, assignments)
4. If description or dueDateTime is provided, it fetches the task details and updates them separately
5. All operations use the `If-Match` header with the ETag to prevent conflicts

## Logging

The script logs all operations to both stderr and the `planner-mcp.log` file. You can monitor the log file to see detailed information about each step.

## Troubleshooting

### "Not authenticated" error

Make sure you have a valid access token:
- Check that `.access-token.txt` exists in the project root
- Or set the `GRAPH_ACCESS_TOKEN` environment variable

### "Invalid JSON in assignments parameter"

Make sure the assignments parameter is valid JSON:
```bash
--assignments '{"user-id@domain.com":{"@odata.type":"#microsoft.graph.plannerAssignment"}}'
```

### "Graph API Error"

Check the error message for details. Common issues:
- Invalid task ID
- Insufficient permissions
- Token expired
- Network issues

## Running with tsx directly

If you prefer not to use npm scripts, you can run the script directly with tsx (no `--` separator needed):

```bash
npx tsx src/test-update-task.ts PlAzDr6l8EGFoo-9yhCgL8kAFyHr --description "test"
```
