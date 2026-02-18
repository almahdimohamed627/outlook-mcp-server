# Outlook Mail MCP Server

A Model Context Protocol (MCP) server that provides email management via Microsoft Graph API.

## Purpose

This MCP server enables AI assistants to manage Outlook emails through Microsoft Graph, providing read, compose, forward, reply, and folder management capabilities.

## Features

### Email Reading
- **`read_emails`** - List emails with filters, sorting, pagination
- **`get_email`** - Get full details of a specific email
- **`get_unread_emails`** - Get all unread emails from inbox
- **`get_draft_emails`** - Get all draft emails
- **`search_emails`** - Search emails using OData filters

### Email Composition
- **`create_draft`** - Create a new email draft
- **`send_email`** - Send an email directly
- **`send_draft`** - Send an existing draft

### Email Actions
- **`forward_email`** - Forward an existing email
- **`reply_email`** - Reply to an email
- **`create_draft_reply`** - Create a draft reply
- **`create_draft_forward`** - Create a draft forward

### Email Management
- **`update_email`** - Update email properties (read status, importance)
- **`mark_as_read`** - Mark email as read
- **`mark_as_unread`** - Mark email as unread
- **`delete_email`** - Delete email (moves to deleted items)
- **`permanent_delete_email`** - Permanently delete email
- **`move_email`** - Move email to folder
- **`copy_email`** - Copy email to folder

### Attachments
- **`list_attachments`** - List all attachments on an email
- **`add_attachment`** - Add attachment to draft

### Folders
- **`list_folders`** - List all email folders
- **`get_mail_folders`** - Get folder details with counts

## Prerequisites

- Docker Desktop with MCP Toolkit enabled
- Docker MCP CLI plugin
- Microsoft 365 account with app registration
- Azure AD app with Mail.Read, Mail.Send permissions

## Installation

### Step 1: Save the Files

```bash
mkdir outlook-mcp-server
cd outlook-mcp-server
# Save all 5 files in this directory
```

### Step 2: Build Docker Image

```bash
docker build -t outlook-mcp-server .
```

### Step 3: Set Up Secrets

Register an app in Azure AD (https://portal.azure.com/#view/Microsoft_AAD_IAM/RegisteredApps):

1. Register new application
2. Add client secret
3. Grant API permissions: Mail.Read, Mail.Send, Mail.ReadWrite
4. Admin consent for permissions

```bash
docker mcp secret set TENANT_ID="your-tenant-id"
docker mcp secret set CLIENT_ID="your-client-id"
docker mcp secret set CLIENT_SECRET="your-client-secret"

docker mcp secret list
```

### Step 4: Create Custom Catalog

```bash
mkdir -p ~/.docker/mcp/catalogs

nano ~/.docker/mcp/catalogs/custom.yaml
```

Add:

```yaml
version: 2
name: custom
displayName: Custom MCP Servers
registry:
  outlook:
    description: "Manage Outlook emails via Microsoft Graph API"
    title: "Outlook Mail"
    type: server
    dateAdded: "2025-08-26T00:00:00Z"
    image: outlook-mcp-server:latest
    ref: ""
    readme: ""
    toolsUrl: ""
    source: ""
    upstream: ""
    icon: ""
    tools:
      - name: read_emails
      - name: get_email
      - name: create_draft
      - name: send_email
      - name: send_draft
      - name: forward_email
      - name: reply_email
      - name: create_draft_reply
      - name: create_draft_forward
      - name: delete_email
      - name: permanent_delete_email
      - name: move_email
      - name: copy_email
      - name: update_email
      - name: list_folders
      - name: list_attachments
      - name: add_attachment
      - name: get_mail_folders
      - name: search_emails
      - name: get_unread_emails
      - name: get_draft_emails
      - name: mark_as_read
      - name: mark_as_unread
    secrets:
      - name: TENANT_ID
        env: TENANT_ID
        example: "xxxxx-xxxxx-xxxxx"
      - name: CLIENT_ID
        env: CLIENT_ID
        example: "xxxxx-xxxxx-xxxxx"
      - name: CLIENT_SECRET
        env: CLIENT_SECRET
        example: "your-secret"
    metadata:
      category: productivity
      tags:
        - email
        - outlook
        - microsoft
        - graph
      license: MIT
      owner: local
```

### Step 5: Update Registry

```bash
nano ~/.docker/mcp/registry.yaml
```

Add under `registry:` key:

```yaml
  outlook:
    ref: ""
```

### Step 6: Configure Claude Desktop

Edit config: `~/Library/Application Support/Claude/claude_desktop_config.json`

```json
{
  "mcpServers": {
    "mcp-toolkit-gateway": {
      "command": "docker",
      "args": [
        "run",
        "-i",
        "--rm",
        "-v", "/var/run/docker.sock:/var/run/docker.sock",
        "-v", "~/.docker/mcp:/mcp",
        "docker/mcp-gateway",
        "--catalog=/mcp/catalogs/docker-mcp.yaml",
        "--catalog=/mcp/catalogs/custom.yaml",
        "--config=/mcp/config.yaml",
        "--registry=/mcp/registry.yaml",
        "--tools-config=/mcp/tools.yaml",
        "--transport=stdio"
      ]
    }
  }
}
```

### Step 7: Restart Claude Desktop

### Step 8: Test

```bash
docker mcp server list
```

## Usage Examples

### Read recent emails

```
"Show me my recent emails"
```

### Search emails

```
"Search for emails about project deadline"
```

### Send email

```
"Send an email to john@example.com about the meeting"
```

### Forward email

```
"Forward the email with ID xyz to jane@example.com"
```

### Mark as read

```
"Mark email abc123 as read"
```

## Filter Examples

The `read_emails` tool supports OData filters:

- Read unread only: `isRead eq false`
- From specific sender: `from/emailAddress/address eq 'sender@example.com'`
- Has attachments: `hasAttachments eq true`
- After date: `receivedDateTime ge 2024-01-01
- Subject contains: `contains(subject,'keyword')`

Combine with AND/OR: `isRead eq false AND hasAttachments eq true`

## Architecture

```
Claude Desktop → MCP Gateway → Outlook MCP Server → Microsoft Graph API
↓
Docker Desktop Secrets
(TENANT_ID, CLIENT_ID, CLIENT_SECRET)
```

## Troubleshooting

### Tools Not Appearing
- Verify Docker image built
- Check catalog and registry files
- Restart Claude Desktop

### Authentication Errors
- Verify secrets: `docker mcp secret list`
- Ensure Azure AD app has correct permissions
- Check admin consent granted

### API Errors
- Verify app has Mail.Read, Mail.Send permissions
- Check tenant ID is correct
- Ensure client secret hasn't expired

## License

MIT License
