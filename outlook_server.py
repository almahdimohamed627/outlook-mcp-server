#!/usr/bin/env python3

"""
Outlook Mail MCP Server - Provides email management via Microsoft Graph API
"""

import os
import sys
import logging
import json
from datetime import datetime
import httpx
from dateutil import parser as date_parser
from mcp.server.fastmcp import FastMCP

# Configure logging to stderr
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    stream=sys.stderr
)
logger = logging.getLogger("outlook-server")

# Initialize MCP server
mcp = FastMCP("outlook")

# Configuration from environment
TENANT_ID = os.environ.get("TENANT_ID", "")
CLIENT_ID = os.environ.get("CLIENT_ID", "")
CLIENT_SECRET = os.environ.get("CLIENT_SECRET", "")
GRAPH_API_VERSION = "v1.0"

# Microsoft Graph base URL
GRAPH_BASE_URL = f"https://graph.microsoft.com/{GRAPH_API_VERSION}"

# Token cache
_access_token = None
_token_expiry = None


async def get_access_token() -> str:
    """Get OAuth2 access token using client credentials flow."""
    global _access_token, _token_expiry
    
    if _access_token and _token_expiry and datetime.now().timestamp() < _token_expiry:
        return _access_token
    
    if not TENANT_ID or not CLIENT_ID or not CLIENT_SECRET:
        raise Exception("Missing required credentials: TENANT_ID, CLIENT_ID, CLIENT_SECRET")
    
    token_url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    
    data = {
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "scope": "https://graph.microsoft.com/.default",
        "grant_type": "client_credentials"
    }
    
    async with httpx.AsyncClient() as client:
        response = await client.post(token_url, data=data, timeout=30)
        response.raise_for_status()
        token_data = response.json()
        
        _access_token = token_data["access_token"]
        expires_in = token_data.get("expires_in", 3600)
        _token_expiry = datetime.now().timestamp() + expires_in - 60
        
        return _access_token


async def make_graph_request(method: str, endpoint: str, data: dict = None, params: dict = None) -> dict:
    """Make authenticated request to Microsoft Graph API."""
    token = await get_access_token()
    url = f"{GRAPH_BASE_URL}{endpoint}"
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }
    
    async with httpx.AsyncClient() as client:
        try:
            if method.upper() == "GET":
                response = await client.get(url, headers=headers, params=params, timeout=30)
            elif method.upper() == "POST":
                response = await client.post(url, headers=headers, json=data, params=params, timeout=30)
            elif method.upper() == "PATCH":
                response = await client.patch(url, headers=headers, json=data, params=params, timeout=30)
            elif method.upper() == "DELETE":
                response = await client.delete(url, headers=headers, timeout=30)
            else:
                raise ValueError(f"Unsupported method: {method}")
            
            response.raise_for_status()
            
            if response.text:
                return response.json()
            return {}
        except httpx.HTTPStatusError as e:
            error_body = e.response.text
            raise Exception(f"Graph API error {e.response.status_code}: {error_body}")


def format_email(msg: dict) -> str:
    """Format email message for display."""
    fields = []
    
    fields.append(f"📧 Subject: {msg.get('subject', 'No Subject')}")
    
    sender = msg.get('from', {})
    sender_email = sender.get('emailAddress', {}).get('address', 'N/A')
    sender_name = sender.get('emailAddress', {}).get('name', 'N/A')
    fields.append(f"👤 From: {sender_name} <{sender_email}>")
    
    to_recipients = msg.get('toRecipients', [])
    to_list = ", ".join([r.get('emailAddress', {}).get('address', '') for r in to_recipients])
    fields.append(f"📤 To: {to_list or 'N/A'}")
    
    cc_recipients = msg.get('ccRecipients', [])
    if cc_recipients:
        cc_list = ", ".join([r.get('emailAddress', {}).get('address', '') for r in cc_recipients])
        fields.append(f"📧 CC: {cc_list}")
    
    fields.append(f"📅 Received: {msg.get('receivedDateTime', 'N/A')}")
    fields.append(f"📤 Sent: {msg.get('sentDateTime', 'N/A')}")
    
    importance = msg.get('importance', 'normal')
    fields.append(f"⭐ Importance: {importance}")
    
    fields.append(f"📎 Has Attachments: {msg.get('hasAttachments', False)}")
    fields.append(f"📝 Is Draft: {msg.get('isDraft', False)}")
    fields.append(f"📖 Is Read: {msg.get('isRead', False)}")
    
    body_preview = msg.get('bodyPreview', '')
    if body_preview:
        fields.append(f"\n📄 Preview:\n{body_preview[:500]}")
    
    msg_id = msg.get('id', '')
    if msg_id:
        fields.append(f"\n🔑 ID: {msg_id}")
    
    return "\n".join(fields)


# === MCP TOOLS ===

@mcp.tool()
async def read_emails(
    folder_id: str = "",
    filter_str: str = "",
    search: str = "",
    select_fields: str = "id,subject,from,toRecipients,ccRecipients,receivedDateTime,sentDateTime,importance,hasAttachments,isDraft,isRead,bodyPreview",
    order_by: str = "receivedDateTime DESC",
    top: str = "25",
    skip: str = "0"
) -> str:
    """List emails from inbox or specified folder with filtering options."""
    try:
        endpoint = "/me/messages"
        
        if folder_id.strip():
            endpoint = f"/me/mailFolders/{folder_id}/messages"
        
        params = {
            "$top": top,
            "$skip": skip,
            "$orderby": order_by
        }
        
        if filter_str.strip():
            params["$filter"] = filter_str
        
        if search.strip():
            params["$search"] = f'"{search}"'
        
        if select_fields.strip():
            params["$select"] = select_fields
        
        result = await make_graph_request("GET", endpoint, params=params)
        
        messages = result.get("value", [])
        
        if not messages:
            return "✅ No emails found"
        
        output = [f"📬 Found {len(messages)} emails:\n"]
        
        for i, msg in enumerate(messages, 1):
            output.append(f"\n--- Email {i} ---")
            output.append(format_email(msg))
        
        # Add pagination info
        if "@odata.nextLink" in result:
            output.append("\n\n⚠️ More emails available. Use skip parameter to paginate.")
        
        return "\n".join(output)
        
    except Exception as e:
        logger.error(f"Error reading emails: {e}")
        return f"❌ Error: {str(e)}"


@mcp.tool()
async def get_email(
    message_id: str = "",
    select_fields: str = "id,subject,from,toRecipients,ccRecipients,bccRecipients,receivedDateTime,sentDateTime,importance,hasAttachments,isDraft,isRead,body,bodyPreview,replyTo,internetMessageHeaders"
) -> str:
    """Get a specific email by ID with full details."""
    try:
        if not message_id.strip():
            return "❌ Error: message_id is required"
        
        params = {}
        if select_fields.strip():
            params["$select"] = select_fields
        
        result = await make_graph_request("GET", f"/me/messages/{message_id}", params=params)
        
        return f"📧 Email Details:\n\n{format_email(result)}"
        
    except Exception as e:
        logger.error(f"Error getting email: {e}")
        return f"❌ Error: {str(e)}"


@mcp.tool()
async def create_draft(
    subject: str = "",
    body: str = "",
    body_type: str = "text",
    to_recipients: str = "",
    cc_recipients: str = "",
    bcc_recipients: str = "",
    importance: str = "normal",
    save_to_sent: str = "false"
) -> str:
    """Create a new email draft."""
    try:
        if not subject.strip() and not body.strip():
            return "❌ Error: Subject or body is required"
        
        message = {
            "subject": subject,
            "importance": importance,
            "body": {
                "contentType": body_type.lower(),
                "content": body
            }
        }
        
        if to_recipients.strip():
            message["toRecipients"] = [
                {"emailAddress": {"address": addr.strip()}} 
                for addr in to_recipients.split(",") if addr.strip()
            ]
        
        if cc_recipients.strip():
            message["ccRecipients"] = [
                {"emailAddress": {"address": addr.strip()}} 
                for addr in cc_recipients.split(",") if addr.strip()
            ]
        
        if bcc_recipients.strip():
            message["bccRecipients"] = [
                {"emailAddress": {"address": addr.strip()}} 
                for addr in bcc_recipients.split(",") if addr.strip()
            ]
        
        result = await make_graph_request("POST", "/me/messages", data=message)
        
        msg_id = result.get("id", "N/A")
        
        return f"✅ Draft created successfully!\n\n📧 Draft ID: {msg_id}\n📝 Subject: {subject}\n📤 To: {to_recipients}"
        
    except Exception as e:
        logger.error(f"Error creating draft: {e}")
        return f"❌ Error: {str(e)}"


@mcp.tool()
async def send_email(
    subject: str = "",
    body: str = "",
    body_type: str = "text",
    to_recipients: str = "",
    cc_recipients: str = "",
    bcc_recipients: str = "",
    importance: str = "normal",
    save_to_sent: str = "true"
) -> str:
    """Send an email directly."""
    try:
        if not to_recipients.strip():
            return "❌ Error: At least one recipient is required"
        
        if not subject.strip() and not body.strip():
            return "❌ Error: Subject or body is required"
        
        message = {
            "subject": subject,
            "importance": importance,
            "saveToSentItems": save_to_sent.lower() == "true",
            "body": {
                "contentType": body_type.lower(),
                "content": body
            }
        }
        
        if to_recipients.strip():
            message["toRecipients"] = [
                {"emailAddress": {"address": addr.strip()}} 
                for addr in to_recipients.split(",") if addr.strip()
            ]
        
        if cc_recipients.strip():
            message["ccRecipients"] = [
                {"emailAddress": {"address": addr.strip()}} 
                for addr in cc_recipients.split(",") if addr.strip()
            ]
        
        if bcc_recipients.strip():
            message["bccRecipients"] = [
                {"emailAddress": {"address": addr.strip()}} 
                for addr in bcc_recipients.split(",") if addr.strip()
            ]
        
        result = await make_graph_request("POST", "/me/sendMail", data={"message": message})
        
        return f"✅ Email sent successfully!\n\n📝 Subject: {subject}\n📤 To: {to_recipients}"
        
    except Exception as e:
        logger.error(f"Error sending email: {e}")
        return f"❌ Error: {str(e)}"


@mcp.tool()
async def send_draft(
    message_id: str = ""
) -> str:
    """Send an existing draft by ID."""
    try:
        if not message_id.strip():
            return "❌ Error: message_id is required"
        
        await make_graph_request("POST", f"/me/messages/{message_id}/send", data={})
        
        return f"✅ Draft sent successfully!\n\n📧 Message ID: {message_id}"
        
    except Exception as e:
        logger.error(f"Error sending draft: {e}")
        return f"❌ Error: {str(e)}"


@mcp.tool()
async def forward_email(
    message_id: str = "",
    to_recipients: str = "",
    comment: str = ""
) -> str:
    """Forward an existing email."""
    try:
        if not message_id.strip():
            return "❌ Error: message_id is required"
        
        if not to_recipients.strip():
            return "❌ Error: to_recipients is required"
        
        forward_data = {
            "toRecipients": [
                {"emailAddress": {"address": addr.strip()}} 
                for addr in to_recipients.split(",") if addr.strip()
            ]
        }
        
        if comment.strip():
            forward_data["comment"] = comment
        
        await make_graph_request("POST", f"/me/messages/{message_id}/forward", data=forward_data)
        
        return f"✅ Email forwarded successfully!\n\n📧 Original Message ID: {message_id}\n📤 Forwarded to: {to_recipients}"
        
    except Exception as e:
        logger.error(f"Error forwarding email: {e}")
        return f"❌ Error: {str(e)}"


@mcp.tool()
async def reply_email(
    message_id: str = "",
    body: str = "",
    reply_all: str = "false"
) -> str:
    """Reply to an email."""
    try:
        if not message_id.strip():
            return "❌ Error: message_id is required"
        
        endpoint = f"/me/messages/{message_id}/reply"
        
        reply_data = {}
        if body.strip():
            reply_data["message"] = {"body": {"contentType": "text", "content": body}}
        
        if reply_all.lower() == "true":
            endpoint = f"/me/messages/{message_id}/replyAll"
        
        await make_graph_request("POST", endpoint, data=reply_data)
        
        reply_type = "reply-all" if reply_all.lower() == "true" else "reply"
        return f"✅ Email {reply_type} sent successfully!\n\n📧 Original Message ID: {message_id}"
        
    except Exception as e:
        logger.error(f"Error replying to email: {e}")
        return f"❌ Error: {str(e)}"


@mcp.tool()
async def create_draft_reply(
    message_id: str = "",
    body: str = "",
    reply_all: str = "false"
) -> str:
    """Create a draft reply to an email."""
    try:
        if not message_id.strip():
            return "❌ Error: message_id is required"
        
        endpoint = f"/me/messages/{message_id}/createReply"
        
        reply_data = {}
        if body.strip():
            reply_data["message"] = {"body": {"contentType": "text", "content": body}}
        
        if reply_all.lower() == "true":
            endpoint = f"/me/messages/{message_id}/createReplyAll"
        
        result = await make_graph_request("POST", endpoint, data=reply_data)
        
        msg_id = result.get("id", "N/A")
        
        return f"✅ Draft reply created!\n\n📧 Draft ID: {msg_id}\n📧 Original Message ID: {message_id}"
        
    except Exception as e:
        logger.error(f"Error creating draft reply: {e}")
        return f"❌ Error: {str(e)}"


@mcp.tool()
async def create_draft_forward(
    message_id: str = "",
    body: str = ""
) -> str:
    """Create a draft forward of an email."""
    try:
        if not message_id.strip():
            return "❌ Error: message_id is required"
        
        forward_data = {}
        if body.strip():
            forward_data["message"] = {"body": {"contentType": "text", "content": body}}
        
        result = await make_graph_request("POST", f"/me/messages/{message_id}/createForward", data=forward_data)
        
        msg_id = result.get("id", "N/A")
        
        return f"✅ Draft forward created!\n\n📧 Draft ID: {msg_id}\n📧 Original Message ID: {message_id}"
        
    except Exception as e:
        logger.error(f"Error creating draft forward: {e}")
        return f"❌ Error: {str(e)}"


@mcp.tool()
async def delete_email(
    message_id: str = ""
) -> str:
    """Delete an email (moves to deleted items)."""
    try:
        if not message_id.strip():
            return "❌ Error: message_id is required"
        
        await make_graph_request("DELETE", f"/me/messages/{message_id}")
        
        return f"✅ Email deleted successfully!\n\n📧 Message ID: {message_id}"
        
    except Exception as e:
        logger.error(f"Error deleting email: {e}")
        return f"❌ Error: {str(e)}"


@mcp.tool()
async def permanent_delete_email(
    message_id: str = ""
) -> str:
    """Permanently delete an email."""
    try:
        if not message_id.strip():
            return "❌ Error: message_id is required"
        
        await make_graph_request("DELETE", f"/me/messages/{message_id}/permanentDelete")
        
        return f"✅ Email permanently deleted!\n\n📧 Message ID: {message_id}"
        
    except Exception as e:
        logger.error(f"Error permanently deleting email: {e}")
        return f"❌ Error: {str(e)}"


@mcp.tool()
async def move_email(
    message_id: str = "",
    destination_folder_id: str = ""
) -> str:
    """Move an email to a different folder."""
    try:
        if not message_id.strip():
            return "❌ Error: message_id is required"
        
        if not destination_folder_id.strip():
            return "❌ Error: destination_folder_id is required"
        
        move_data = {
            "destinationId": destination_folder_id
        }
        
        result = await make_graph_request("POST", f"/me/messages/{message_id}/move", data=move_data)
        
        new_id = result.get("id", "N/A")
        
        return f"✅ Email moved successfully!\n\n📧 Original ID: {message_id}\n📁 New ID: {new_id}\n📁 Destination Folder: {destination_folder_id}"
        
    except Exception as e:
        logger.error(f"Error moving email: {e}")
        return f"❌ Error: {str(e)}"


@mcp.tool()
async def copy_email(
    message_id: str = "",
    destination_folder_id: str = ""
) -> str:
    """Copy an email to a different folder."""
    try:
        if not message_id.strip():
            return "❌ Error: message_id is required"
        
        if not destination_folder_id.strip():
            return "❌ Error: destination_folder_id is required"
        
        copy_data = {
            "destinationId": destination_folder_id
        }
        
        result = await make_graph_request("POST", f"/me/messages/{message_id}/copy", data=copy_data)
        
        new_id = result.get("id", "N/A")
        
        return f"✅ Email copied successfully!\n\n📧 Original ID: {message_id}\n📁 Copy ID: {new_id}\n📁 Destination Folder: {destination_folder_id}"
        
    except Exception as e:
        logger.error(f"Error copying email: {e}")
        return f"❌ Error: {str(e)}"


@mcp.tool()
async def update_email(
    message_id: str = "",
    is_read: str = "",
    is_flagged: str = "",
    importance: str = "",
    subject: str = ""
) -> str:
    """Update email properties (read status, importance, subject, etc.)."""
    try:
        if not message_id.strip():
            return "❌ Error: message_id is required"
        
        update_data = {}
        
        if is_read.strip():
            update_data["isRead"] = is_read.strip().lower() == "true"
        
        if is_flagged.strip():
            update_data["isFlagged"] = is_flagged.strip().lower() == "true"
        
        if importance.strip():
            update_data["importance"] = importance.strip().lower()
        
        if subject.strip():
            update_data["subject"] = subject
        
        if not update_data:
            return "❌ Error: No properties to update"
        
        result = await make_graph_request("PATCH", f"/me/messages/{message_id}", data=update_data)
        
        return f"✅ Email updated successfully!\n\n📧 Message ID: {message_id}\n📝 Updates: {json.dumps(update_data, indent=2)}"
        
    except Exception as e:
        logger.error(f"Error updating email: {e}")
        return f"❌ Error: {str(e)}"


@mcp.tool()
async def list_folders() -> str:
    """List all email folders in the mailbox."""
    try:
        result = await make_graph_request("GET", "/me/mailFolders")
        
        folders = result.get("value", [])
        
        if not folders:
            return "✅ No folders found"
        
        output = [f"📁 Found {len(folders)} folders:\n"]
        
        for folder in folders:
            output.append(f"\n📁 {folder.get('displayName', 'N/A')}")
            output.append(f"   ID: {folder.get('id', 'N/A')}")
            output.append(f"   Total Items: {folder.get('totalItemCount', 0)}")
            output.append(f"   Unread: {folder.get('unreadItemCount', 0)}")
        
        return "\n".join(output)
        
    except Exception as e:
        logger.error(f"Error listing folders: {e}")
        return f"❌ Error: {str(e)}"


@mcp.tool()
async def list_attachments(
    message_id: str = ""
) -> str:
    """List all attachments for an email."""
    try:
        if not message_id.strip():
            return "❌ Error: message_id is required"
        
        result = await make_graph_request("GET", f"/me/messages/{message_id}/attachments")
        
        attachments = result.get("value", [])
        
        if not attachments:
            return "✅ No attachments found"
        
        output = [f"📎 Found {len(attachments)} attachments:\n"]
        
        for att in attachments:
            att_type = att.get("@odata.type", "unknown")
            name = att.get("name", "N/A")
            size = att.get("size", 0)
            
            output.append(f"\n📎 {name}")
            output.append(f"   Type: {att_type}")
            output.append(f"   Size: {size} bytes")
        
        return "\n".join(output)
        
    except Exception as e:
        logger.error(f"Error listing attachments: {e}")
        return f"❌ Error: {str(e)}"


@mcp.tool()
async def add_attachment(
    message_id: str = "",
    file_name: str = "",
    content_bytes: str = ""
) -> str:
    """Add an attachment to an email draft."""
    try:
        if not message_id.strip():
            return "❌ Error: message_id is required"
        
        if not file_name.strip():
            return "❌ Error: file_name is required"
        
        if not content_bytes.strip():
            return "❌ Error: content_bytes is required"
        
        import base64
        
        attachment_data = {
            "@odata.type": "#microsoft.graph.fileAttachment",
            "name": file_name,
            "contentBytes": content_bytes
        }
        
        result = await make_graph_request("POST", f"/me/messages/{message_id}/attachments", data=attachment_data)
        
        att_id = result.get("id", "N/A")
        
        return f"✅ Attachment added!\n\n📎 Attachment ID: {att_id}\n📄 File name: {file_name}"
        
    except Exception as e:
        logger.error(f"Error adding attachment: {e}")
        return f"❌ Error: {str(e)}"


@mcp.tool()
async def get_mail_folders(
    top: str = "100"
) -> str:
    """Get top mail folders with item counts."""
    try:
        params = {
            "$top": top,
            "$select": "id,displayName,totalItemCount,unreadItemCount,childFolderCount"
        }
        
        result = await make_graph_request("GET", "/me/mailFolders", params=params)
        
        folders = result.get("value", [])
        
        if not folders:
            return "✅ No folders found"
        
        output = [f"📁 Mail Folders:\n"]
        
        for folder in folders:
            name = folder.get("displayName", "N/A")
            total = folder.get("totalItemCount", 0)
            unread = folder.get("unreadItemCount", 0)
            children = folder.get("childFolderCount", 0)
            
            output.append(f"\n📁 {name}")
            output.append(f"   Total: {total} | Unread: {unread} | Subfolders: {children}")
        
        return "\n".join(output)
        
    except Exception as e:
        logger.error(f"Error getting mail folders: {e}")
        return f"❌ Error: {str(e)}"


@mcp.tool()
async def search_emails(
    query: str = "",
    filter_str: str = "",
    top: str = "25"
) -> str:
    """Search emails using Microsoft Search or OData filters."""
    try:
        if not query.strip() and not filter_str.strip():
            return "❌ Error: Either query or filter_str is required"
        
        params = {
            "$top": top,
            "$select": "id,subject,from,toRecipients,receivedDateTime,importance,hasAttachments,isRead"
        }
        
        if query.strip():
            params["$search"] = f'"{query}"'
        
        if filter_str.strip():
            params["$filter"] = filter_str
        
        result = await make_graph_request("GET", "/me/messages", params=params)
        
        messages = result.get("value", [])
        
        if not messages:
            return "✅ No emails found matching criteria"
        
        output = [f"🔍 Found {len(messages)} emails:\n"]
        
        for msg in messages:
            subject = msg.get("subject", "No Subject")
            from_addr = msg.get("from", {}).get("emailAddress", {}).get("address", "N/A")
            received = msg.get("receivedDateTime", "N/A")
            is_read = msg.get("isRead", False)
            
            read_status = "📖" if is_read else "📕"
            output.append(f"\n{read_status} {subject}")
            output.append(f"   From: {from_addr}")
            output.append(f"   Received: {received}")
        
        return "\n".join(output)
        
    except Exception as e:
        logger.error(f"Error searching emails: {e}")
        return f"❌ Error: {str(e)}"


@mcp.tool()
async def get_unread_emails(
    top: str = "25"
) -> str:
    """Get all unread emails from inbox."""
    try:
        params = {
            "$filter": "isRead eq false",
            "$top": top,
            "$orderby": "receivedDateTime DESC",
            "$select": "id,subject,from,toRecipients,receivedDateTime,importance,hasAttachments"
        }
        
        result = await make_graph_request("GET", "/me/messages", params=params)
        
        messages = result.get("value", [])
        
        if not messages:
            return "✅ No unread emails"
        
        output = [f"📕 Found {len(messages)} unread emails:\n"]
        
        for msg in messages:
            output.append(f"\n--- Unread Email ---")
            output.append(format_email(msg))
        
        return "\n".join(output)
        
    except Exception as e:
        logger.error(f"Error getting unread emails: {e}")
        return f"❌ Error: {str(e)}"


@mcp.tool()
async def get_draft_emails(
    top: str = "25"
) -> str:
    """Get all draft emails."""
    try:
        params = {
            "$filter": "isDraft eq true",
            "$top": top,
            "$orderby": "createdDateTime DESC",
            "$select": "id,subject,toRecipients,ccRecipients,createdDateTime"
        }
        
        result = await make_graph_request("GET", "/me/messages", params=params)
        
        messages = result.get("value", [])
        
        if not messages:
            return "✅ No draft emails found"
        
        output = [f"📝 Found {len(messages)} draft emails:\n"]
        
        for msg in messages:
            subject = msg.get("subject", "No Subject")
            created = msg.get("createdDateTime", "N/A")
            msg_id = msg.get("id", "N/A")
            
            to_recipients = msg.get("toRecipients", [])
            to_list = ", ".join([r.get("emailAddress", {}).get("address", "") for r in to_recipients])
            
            output.append(f"\n📝 Draft: {subject}")
            output.append(f"   To: {to_list or 'N/A'}")
            output.append(f"   Created: {created}")
            output.append(f"   ID: {msg_id}")
        
        return "\n".join(output)
        
    except Exception as e:
        logger.error(f"Error getting draft emails: {e}")
        return f"❌ Error: {str(e)}"


@mcp.tool()
async def mark_as_read(
    message_id: str = ""
) -> str:
    """Mark an email as read."""
    try:
        if not message_id.strip():
            return "❌ Error: message_id is required"
        
        update_data = {"isRead": True}
        await make_graph_request("PATCH", f"/me/messages/{message_id}", data=update_data)
        
        return f"✅ Email marked as read!\n\n📧 Message ID: {message_id}"
        
    except Exception as e:
        logger.error(f"Error marking as read: {e}")
        return f"❌ Error: {str(e)}"


@mcp.tool()
async def mark_as_unread(
    message_id: str = ""
) -> str:
    """Mark an email as unread."""
    try:
        if not message_id.strip():
            return "❌ Error: message_id is required"
        
        update_data = {"isRead": False}
        await make_graph_request("PATCH", f"/me/messages/{message_id}", data=update_data)
        
        return f"✅ Email marked as unread!\n\n📧 Message ID: {message_id}"
        
    except Exception as e:
        logger.error(f"Error marking as unread: {e}")
        return f"❌ Error: {str(e)}"


# === SERVER STARTUP ===

if __name__ == "__main__":
    logger.info("Starting Outlook Mail MCP server...")
    
    if not TENANT_ID:
        logger.warning("TENANT_ID not set - set via environment variable")
    if not CLIENT_ID:
        logger.warning("CLIENT_ID not set - set via environment variable")
    if not CLIENT_SECRET:
        logger.warning("CLIENT_SECRET not set - set via environment variable")
    
    try:
        mcp.run(transport='stdio')
    except Exception as e:
        logger.error(f"Server error: {e}", exc_info=True)
        sys.exit(1)
