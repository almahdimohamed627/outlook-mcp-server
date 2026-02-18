# Outlook MCP Server Implementation Details

## Overview

This MCP server integrates with Microsoft Graph API to provide comprehensive Outlook email management. It uses OAuth2 client credentials flow for authentication.

## Authentication

The server uses Azure AD OAuth2 client credentials flow:

1. **TENANT_ID** - Azure AD tenant ID
2. **CLIENT_ID** - Registered application client ID  
3. **CLIENT_SECRET** - Application client secret

Token is automatically cached and refreshed. Required permissions:
- `Mail.Read` - Read emails
- `Mail.Send` - Send emails
- `Mail.ReadWrite` - Full mail access
- `MailboxSettings.Read` - Read mailbox settings

## API Endpoints Used

All operations use Microsoft Graph v1.0:
- `GET /me/messages` - List messages
- `GET /me/messages/{id}` - Get message
- `POST /me/messages` - Create draft
- `PATCH /me/messages/{id}` - Update message
- `DELETE /me/messages/{id}` - Delete message
- `POST /me/messages/{id}/forward` - Forward
- `POST /me/messages/{id}/reply` - Reply
- `POST /me/messages/{id}/send` - Send draft
- `POST /me/sendMail` - Send directly
- `GET /me/mailFolders` - List folders
- `POST /me/messages/{id}/move` - Move message
- `POST /me/messages/{id}/copy` - Copy message

## Tool Parameters

All parameters use string types with empty string defaults for MCP compatibility:

| Parameter | Type | Description |
|-----------|------|-------------|
| message_id | str | Email ID (GUID) |
| folder_id | str | Mail folder ID |
| subject | str | Email subject |
| body | str | Email body content |
| body_type | str | "text" or "html" |
| to_recipients | str | Comma-separated emails |
| cc_recipients | str | Comma-separated emails |
| bcc_recipients | str | Comma-separated emails |
| importance | str | "low", "normal", "high" |
| filter_str | str | OData filter expression |
| search | str | Search query |
| top | str | Max results (number as string) |
| skip | str | Pagination offset |
| order_by | str | OData orderby clause |
| select_fields | str | Comma-separated fields |
| is_read | str | "true" or "false" |
| is_flagged | str | "true" or "false" |
| destination_folder_id | str | Target folder ID |
| comment | str | Forward/reply comment |
| file_name | str | Attachment file name |
| content_bytes | str | Base64 encoded content |

## Error Handling

All tools return formatted strings:
- Success: `✅ Success message`
- Error: `❌ Error: description`

Errors are logged to stderr with full context.

## Response Formatting

Email data formatted with emojis:
- 📧 - Email/message
- 📬 - Multiple emails
- 📕 - Unread
- 📖 - Read
- 📝 - Draft
- 📁 - Folder
- 📎 - Attachment
- ⭐ - Importance
- 👤 - Sender
- 📤 - Recipient
- 📅 - Date/time
- 🔑 - ID

## Common OData Filters

```
isRead eq false                    # Unread only
hasAttachments eq true            # Has attachments
from/emailAddress/address eq 'x'   # From specific sender
contains(subject,'keyword')        # Subject contains
receivedDateTime ge 2024-01-01     # After date
importance eq 'high'               # High importance
isDraft eq true                    # Drafts only
```

## Development Notes

- Uses httpx async client for HTTP requests
- Token caching with automatic refresh
- All API calls include timeout (30s default)
- Non-root user in Docker for security
- Logging to stderr for container monitoring

## Testing

Test with curl after getting access token:

```bash
curl -H "Authorization: Bearer $TOKEN" \
  "https://graph.microsoft.com/v1.0/me/messages?$top=5"
```
