# Planner MCP Server

A Model Context Protocol (MCP) server that enables Claude to interact with Microsoft Planner through the Microsoft Graph API.

## Features

- **Authentication**: OAuth 2.0 device code flow for secure authentication
- **Plans Management**: List and view Planner plans
- **Buckets Management**: List, create, and manage buckets within plans
- **Tasks Management**: Create, update, delete, and view tasks
- **Token Persistence**: Stores access tokens across server restarts

## Prerequisites

- Node.js 18+ or newer
- npm or yarn
- Microsoft account with access to Planner

## Installation

1. Clone or download this repository
2. Install dependencies:

```bash
cd planner-mcp
npm install
```

3. Build the project:

```bash
npm run build
```

## Authentication

Before using the Planner tools, you need to authenticate:

1. Start the authentication flow by calling the `auth_start` tool
2. Visit the displayed verification URL
3. Enter the provided code
4. Sign in with your Microsoft account
5. Use the `auth_poll` tool to check if authentication is complete

The access token will be saved locally and reused for future sessions.

## Available Tools

### Authentication Tools

- **`auth_start`**: Start the Planner authentication flow (device code)
- **`auth_poll`**: Check if authentication is complete

### Plan Tools

- **`list_plans`**: List all Microsoft Planner plans for the current user
- **`get_plan`**: Get detailed information about a specific plan

### Bucket Tools

- **`list_buckets`**: List all buckets in a plan
- **`create_bucket`**: Create a new bucket in a plan

### Task Tools

- **`list_tasks`**: List tasks in a plan or bucket
- **`get_task`**: Get detailed information about a specific task
- **`create_task`**: Create a new task in a bucket
- **`update_task`**: Update an existing task
- **`delete_task`**: Delete a task

## Usage with Claude Desktop

Add this server to your Claude Desktop configuration file:

**Windows**: `%APPDATA%\Claude\claude_desktop_config.json`
**macOS**: `~/Library/Application Support/Claude/claude_desktop_config.json`

```json
{
  "mcpServers": {
    "planner": {
      "command": "node",
      "args": ["C:\\path\\to\\planner-mcp\\build\\index.js"]
    }
  }
}
```

## Required Permissions

The Microsoft Graph API requires the following permissions for Planner:

- `Group.Read.All` - To list plans and buckets
- `Tasks.ReadWrite` - To read, create, update, and delete tasks

## Example Workflows

### Create a new task

1. List plans: `list_plans`
2. List buckets in a plan: `list_buckets` with the plan ID
3. Create a task: `create_task` with plan ID, bucket ID, and task details

### Move a task to another bucket

1. Get task details: `get_task`
2. Update task: `update_task` with the new bucket ID

### Update task due date

1. Get task details: `get_task`
2. Update task: `update_task` with the new due date in ISO 8601 format (e.g., "2025-12-31T23:59:59Z")

## Troubleshooting

### Authentication Issues

If authentication fails, try these steps:

1. Delete the `.access-token.txt` file
2. Call `auth_start` again
3. Make sure you're using a Microsoft account that has access to Planner

### "Not authenticated" Error

If you get a "Not authenticated" error:

1. Make sure you've completed the authentication flow
2. Check that the `.access-token.txt` file exists and contains a valid token
3. If the token has expired, run `auth_start` again

### Permission Errors

If you get permission errors:

- Make sure your Microsoft account has access to the specified plan
- Check that the required permissions are granted in your Azure AD app

## Development

### Build

```bash
npm run build
```

### Watch mode

```bash
npm run watch
```

### Start server

```bash
npm start
```

## License

ISC

## Contributing

Contributions are welcome! Please feel free to submit issues or pull requests.
