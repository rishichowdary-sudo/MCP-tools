# Multi-Service AI Agent Platform — Product Requirements Document (PRD)

**Version:** 2.0
**Date:** February 20, 2026
**Status:** Living Document

---

## Table of Contents

1. [Executive Summary](#1-executive-summary)
2. [System Architecture](#2-system-architecture)
3. [Technology Stack](#3-technology-stack)
4. [Service Integrations & Tools](#4-service-integrations--tools)
5. [Frontend Widgets & UI](#5-frontend-widgets--ui)
6. [Meeting Transcription](#6-meeting-transcription)
7. [AI Chat Engine](#7-ai-chat-engine)
8. [Authentication & OAuth](#8-authentication--oauth)
9. [API Endpoints](#9-api-endpoints)
10. [Security Architecture](#10-security-architecture)
11. [Configuration & Environment](#11-configuration--environment)
12. [Data Flow & Tool Execution](#12-data-flow--tool-execution)
13. [Scheduled Tasks & Timer](#13-scheduled-tasks--timer)
14. [File Management](#14-file-management)
15. [Performance & Limits](#15-performance--limits)
16. [Dependencies](#16-dependencies)

---

## 1. Executive Summary

The **Multi-Service AI Agent Platform** is a comprehensive AI-powered productivity hub that integrates **12 services** with **155+ tools** into a single conversational interface. Users interact with a natural language chat agent (powered by OpenAI GPT-4) that intelligently routes requests to the appropriate service — whether it's sending an email, creating a calendar event, managing GitHub repos, querying Google Cloud Storage, or summarizing a meeting transcription.

### Key Highlights

| Metric | Value |
|---|---|
| Total Integrated Services | 12 |
| Total Tools | 155+ |
| API Endpoints | 50+ |
| Frontend Widgets | 13 |
| OAuth Providers | 3 (Google, GitHub, Microsoft) |
| AI Model | GPT-4o-mini (primary), GPT-4o (fallback) |

### Integrated Services at a Glance

```
┌─────────────────────────────────────────────────────────────────┐
│                  MULTI-SERVICE AI AGENT PLATFORM                │
├──────────────┬──────────────┬──────────────┬────────────────────┤
│   Google     │   Microsoft  │   GitHub     │   Infrastructure   │
├──────────────┼──────────────┼──────────────┼────────────────────┤
│ Gmail (25)   │ Outlook (16) │ Repos (20)   │ GCS (16)           │
│ Calendar (19)│ Teams (10)   │              │ Timer Tasks        │
│ Drive (15)   │              │              │ Meeting Trans. (3) │
│ Sheets (11)  │              │              │                    │
│ Docs (8)     │              │              │                    │
│ Chat (3)     │              │              │                    │
└──────────────┴──────────────┴──────────────┴────────────────────┘
  (Numbers indicate tool count per service)
```

---

## 2. System Architecture

### High-Level Architecture Diagram

```
┌──────────────────────────────────────────────────────────────────────────┐
│                              CLIENT LAYER                                │
│                                                                          │
│  ┌────────────────────────────────────────────────────────────────────┐  │
│  │                      Web Application (SPA)                         │  │
│  │  ┌──────────┐ ┌──────────┐ ┌──────────┐ ┌──────────┐ ┌────────┐  │  │
│  │  │  Chat    │ │ Service  │ │ Capabil. │ │  File    │ │ Timer  │  │  │
│  │  │  Window  │ │ Panels   │ │  Modal   │ │ Upload   │ │ Panel  │  │  │
│  │  └────┬─────┘ └────┬─────┘ └──────────┘ └────┬─────┘ └───┬────┘  │  │
│  └───────┼────────────┼──────────────────────────┼───────────┼───────┘  │
└──────────┼────────────┼──────────────────────────┼───────────┼──────────┘
           │            │         HTTP/REST         │           │
           ▼            ▼                           ▼           ▼
┌──────────────────────────────────────────────────────────────────────────┐
│                            SERVER LAYER                                  │
│                         (Node.js + Express)                              │
│                                                                          │
│  ┌─────────────────────────────────────────────────────────────────────┐ │
│  │                        Middleware Stack                              │ │
│  │  ┌────────┐  ┌──────────┐  ┌───────────┐  ┌────────────────────┐   │ │
│  │  │  CORS  │  │  Rate    │  │ Security  │  │  Body Parser       │   │ │
│  │  │        │  │  Limiter │  │  Headers  │  │  (5GB limit)       │   │ │
│  │  └────────┘  └──────────┘  └───────────┘  └────────────────────┘   │ │
│  └─────────────────────────────────────────────────────────────────────┘ │
│                                                                          │
│  ┌──────────────────────────┐    ┌──────────────────────────────────┐   │
│  │    AI Chat Engine        │    │      OAuth Manager               │   │
│  │  ┌────────────────────┐  │    │  ┌────────┐ ┌──────┐ ┌───────┐  │   │
│  │  │  OpenAI GPT-4      │  │    │  │ Google │ │GitHub│ │MS     │  │   │
│  │  │  Tool Router       │  │    │  │ OAuth  │ │OAuth │ │OAuth  │  │   │
│  │  │  History Manager   │  │    │  └────────┘ └──────┘ └───────┘  │   │
│  │  └────────────────────┘  │    └──────────────────────────────────┘   │
│  └──────────┬───────────────┘                                           │
│             │  Tool Calls                                                │
│             ▼                                                            │
│  ┌──────────────────────────────────────────────────────────────────┐   │
│  │                    TOOL EXECUTION LAYER                          │   │
│  │                                                                  │   │
│  │  ┌────────┐ ┌────────┐ ┌──────┐ ┌───────┐ ┌──────┐ ┌────────┐  │   │
│  │  │ Gmail  │ │Calendar│ │Drive │ │Sheets │ │ Docs │ │ GChat  │  │   │
│  │  │ (25)   │ │ (19)   │ │(15)  │ │ (11)  │ │ (8)  │ │  (3)   │  │   │
│  │  └────────┘ └────────┘ └──────┘ └───────┘ └──────┘ └────────┘  │   │
│  │  ┌────────┐ ┌────────┐ ┌──────┐ ┌───────┐ ┌────────────────┐   │   │
│  │  │GitHub  │ │Outlook │ │Teams │ │  GCS  │ │   Meeting      │   │   │
│  │  │ (20)   │ │ (16)   │ │(10)  │ │ (16)  │ │   Transcript.  │   │   │
│  │  └────────┘ └────────┘ └──────┘ └───────┘ └────────────────┘   │   │
│  └──────────────────────────────────────────────────────────────────┘   │
│                                                                          │
│  ┌──────────────────────┐   ┌──────────────────────────────────────┐   │
│  │  MCP Bridge          │   │  File Manager                        │   │
│  │  (Sheets/Drive)      │   │  (Upload, Download, Cleanup)         │   │
│  └──────────────────────┘   └──────────────────────────────────────┘   │
└──────────────────────────────────────────────────────────────────────────┘
                                    │
                                    ▼
┌──────────────────────────────────────────────────────────────────────────┐
│                         EXTERNAL SERVICES                                │
│                                                                          │
│  ┌──────────────┐  ┌──────────────┐  ┌──────────────┐  ┌────────────┐  │
│  │  Google APIs  │  │  GitHub API  │  │  MS Graph    │  │  OpenAI    │  │
│  │  (googleapis) │  │  (Octokit)   │  │  API         │  │  API       │  │
│  └──────────────┘  └──────────────┘  └──────────────┘  └────────────┘  │
└──────────────────────────────────────────────────────────────────────────┘
```

### Component Interaction Flow

```
  User Message
       │
       ▼
┌──────────────┐     ┌──────────────┐     ┌────────────────┐
│   Frontend   │────▶│  /api/chat   │────▶│  OpenAI GPT-4  │
│   (app.js)   │     │  endpoint    │     │  with tools    │
└──────────────┘     └──────┬───────┘     └───────┬────────┘
       ▲                    │                     │
       │                    │              Tool Selection
       │                    │                     │
       │                    ▼                     ▼
       │             ┌──────────────┐     ┌────────────────┐
       │             │   Response   │◀────│  Tool Executor │
       │             │   Builder    │     │  (155+ tools)  │
       │             └──────┬───────┘     └────────────────┘
       │                    │
       └────────────────────┘
            JSON Response
```

---

## 3. Technology Stack

### Backend

| Component | Technology | Purpose |
|---|---|---|
| Runtime | Node.js | Server-side JavaScript |
| Framework | Express.js v4.18 | HTTP server & routing |
| AI Engine | OpenAI SDK v4.28 | GPT-4 chat completions & tool calling |
| Google APIs | googleapis v129 | Gmail, Calendar, Drive, Sheets, Docs, Chat |
| GitHub API | Octokit v5 | Repository, issue, PR management |
| Microsoft API | MS Graph (via REST) | Outlook, Teams |
| Cloud Storage | @google-cloud/storage | GCS bucket & object management |
| MCP Client | @modelcontextprotocol/sdk v1.26 | Sheets/Drive MCP bridge |
| File Upload | Multer v2 | Multipart form handling |
| Rate Limiting | express-rate-limit v8 | API abuse prevention |
| WebSocket | ws v8 | Real-time communication |

### Frontend

| Component | Technology | Purpose |
|---|---|---|
| UI Framework | Vanilla HTML5/CSS3/JS | Zero-dependency frontend |
| Styling | Custom CSS | Responsive design with modals |
| State Management | In-memory JS objects | Service connection states |
| Communication | Fetch API | REST calls to backend |

### Storage

| What | Where | Format |
|---|---|---|
| Google OAuth Tokens | `~/.gmail-mcp/token.json` | JSON |
| GitHub Tokens | `~/.gmail-mcp/github-token.json` | JSON |
| Outlook Tokens | `~/.gmail-mcp/outlook-token.json` | JSON |
| Scheduled Tasks | `~/.gmail-mcp/scheduled-tasks.json` | JSON |
| File Uploads | `./uploads/` | Temporary (15-min TTL) |

---

## 4. Service Integrations & Tools

### 4.1 Gmail (25 Tools)

Full email lifecycle management through the Gmail API.

```
Gmail Tools
├── Compose & Send
│   ├── send_email           — Send email with To/CC/BCC, attachments
│   ├── create_draft         — Save email as draft
│   ├── reply_to_email       — Reply to existing thread
│   ├── forward_email        — Forward email to recipients
│   ├── send_draft           — Send a saved draft
│   └── delete_draft         — Remove a draft
│
├── Read & Search
│   ├── search_emails        — Search with Gmail query syntax
│   ├── read_email           — Read full email content
│   ├── list_emails          — List inbox emails
│   ├── get_thread           — Get full conversation thread
│   ├── list_drafts          — List all drafts
│   └── get_attachment_info  — Get attachment metadata
│
├── Organization
│   ├── modify_labels        — Add/remove labels
│   ├── create_label         — Create custom label
│   ├── delete_label         — Delete a label
│   ├── list_labels          — List all labels
│   ├── archive_email        — Archive email
│   ├── trash_email          — Move to trash
│   ├── untrash_email        — Restore from trash
│   ├── star_email           — Star an email
│   ├── unstar_email         — Unstar an email
│   ├── mark_as_read         — Mark as read
│   └── mark_as_unread       — Mark as unread
│
├── Bulk Operations
│   └── batch_modify_emails  — Batch label/archive operations
│
└── Account
    └── get_profile           — Get email profile info
```

### 4.2 Google Calendar (19 Tools)

Full calendar management with availability checking and Google Meet integration.

```
Calendar Tools
├── Event Management
│   ├── list_events              — List events with date range filters
│   ├── get_event                — Get single event details
│   ├── create_event             — Create event with attendees
│   ├── create_meet_event        — Create event with auto-generated Meet link
│   ├── add_meet_link_to_event   — Add Meet link to existing event
│   ├── update_event             — Modify event details
│   ├── delete_event             — Delete an event
│   ├── quick_add_event          — Natural language event creation
│   ├── move_event               — Move event to another calendar
│   └── update_event_attendees   — Add/remove attendees
│
├── Availability
│   ├── get_free_busy            — Check free/busy time blocks
│   ├── check_person_availability— Check a person's availability
│   └── find_common_free_slots   — Find common open times for groups
│
├── Calendar Management
│   ├── list_calendars           — List all calendars
│   ├── create_calendar          — Create a new calendar
│   ├── clear_calendar           — Clear all events from calendar
│   ├── get_calendar_colors      — Get available color options
│   └── watch_events             — Set up event notifications
│
└── Recurring Events
    └── list_recurring_instances  — List instances of recurring event
```

### 4.3 Google Drive (15 Tools)

File storage, sharing, and format conversion capabilities.

```
Drive Tools
├── File Operations
│   ├── list_drive_files              — Browse files and folders
│   ├── get_drive_file                — Get file metadata
│   ├── create_drive_file             — Create new file
│   ├── create_drive_folder           — Create new folder
│   ├── update_drive_file             — Update file content
│   ├── delete_drive_file             — Delete a file
│   ├── copy_drive_file               — Copy a file
│   └── move_drive_file               — Move file between folders
│
├── Download & Extract
│   ├── download_drive_file           — Download file (in-memory)
│   ├── download_drive_file_to_local  — Download to local filesystem
│   └── extract_drive_file_text       — Extract text from any file
│
├── Content
│   └── append_drive_document_text    — Append text to a document
│
├── Sharing
│   └── share_drive_file              — Share with users/permissions
│
└── Conversion
    ├── convert_file_to_google_doc    — Convert uploaded file → Google Doc
    └── convert_file_to_google_sheet  — Convert uploaded file → Google Sheet
```

### 4.4 Google Sheets (11 Tools)

Spreadsheet creation, reading, writing, and tab management.

```
Sheets Tools
├── Spreadsheet Management
│   ├── list_spreadsheets      — List all spreadsheets
│   ├── create_spreadsheet     — Create new spreadsheet
│   └── get_spreadsheet        — Get spreadsheet metadata
│
├── Tab Management
│   ├── list_sheet_tabs        — List tabs in a spreadsheet
│   ├── add_sheet_tab          — Add new tab
│   └── delete_sheet_tab       — Delete a tab
│
└── Data Operations
    ├── read_sheet_values       — Read cell range
    ├── update_sheet_values     — Write to cell range
    ├── update_timesheet_hours  — Update timesheet entries
    ├── append_sheet_values     — Append rows to sheet
    └── clear_sheet_values      — Clear cell range
```

**MCP Bridge:** When `SHEETS_MCP_ENABLED=true`, additional tools from `@isaacphi/mcp-gdrive` are dynamically discovered and available for advanced Sheets/Drive operations.

### 4.5 Google Docs (8 Tools)

Document creation and content manipulation.

```
Docs Tools
├── Document Management
│   ├── list_documents     — List all documents
│   ├── get_document       — Get document structure
│   ├── create_document    — Create new document
│   └── get_document_text  — Extract plain text
│
└── Content Editing
    ├── insert_text        — Insert text at position
    ├── replace_text       — Find and replace text
    ├── delete_content     — Delete content range
    └── append_text        — Append text to end
```

### 4.6 Google Chat (3 Tools)

Google Workspace Chat space and messaging support.

```
Chat Tools
├── list_chat_spaces     — List available spaces
├── send_chat_message    — Send message to space
└── list_chat_messages   — Read messages from space
```

### 4.7 GitHub (20+ Tools)

Full repository, issue, PR, and code management.

```
GitHub Tools
├── Repositories
│   ├── list_repos              — List user repositories
│   ├── get_repo                — Get repo details
│   ├── create_repo             — Create new repository
│   └── search_repos            — Search public repos
│
├── Issues
│   ├── list_issues             — List repo issues
│   ├── create_issue            — Create new issue
│   └── update_issue            — Update issue state/content
│
├── Pull Requests
│   ├── list_pull_requests      — List PRs
│   ├── get_pull_request        — Get PR details
│   ├── create_pull_request     — Open new PR
│   └── merge_pull_request      — Merge a PR
│
├── Branches & Commits
│   ├── list_branches           — List branches
│   ├── create_branch           — Create new branch
│   ├── list_commits            — List commit history
│   ├── revert_commit           — Revert a commit
│   └── reset_branch            — Reset branch to commit
│
├── Code
│   ├── get_file_content        — Read file from repo
│   ├── create_or_update_file   — Write file to repo
│   └── search_code             — Search code across repos
│
└── Account
    ├── get_user_profile        — Get user info
    ├── list_notifications      — List notifications
    └── list_gists              — List user gists
```

### 4.8 Microsoft Outlook (16 Tools)

Full Outlook email management via Microsoft Graph API.

```
Outlook Tools
├── Compose & Send
│   ├── outlook_send_email       — Send email
│   ├── outlook_reply_to_email   — Reply to email
│   ├── outlook_forward_email    — Forward email
│   ├── outlook_create_draft     — Create draft
│   └── outlook_send_draft       — Send draft
│
├── Read & Search
│   ├── outlook_list_emails      — List inbox emails
│   ├── outlook_read_email       — Read email content
│   ├── outlook_search_emails    — Search emails
│   ├── outlook_list_drafts      — List drafts
│   └── outlook_get_attachments  — Get attachments
│
├── Organization
│   ├── outlook_delete_email     — Delete email
│   ├── outlook_move_email       — Move between folders
│   ├── outlook_mark_as_read     — Mark read
│   ├── outlook_mark_as_unread   — Mark unread
│   ├── outlook_flag_email       — Flag email
│   ├── outlook_list_folders     — List mail folders
│   └── outlook_create_folder    — Create folder
│
└── Account
    └── outlook_get_user_profile — Get user profile
```

### 4.9 Microsoft Teams (10 Tools)

Team, channel, and chat management via Microsoft Graph API.

```
Teams Tools
├── Teams & Channels
│   ├── teams_list_teams            — List joined teams
│   ├── teams_get_team              — Get team details
│   ├── teams_list_channels         — List channels
│   ├── teams_send_channel_message  — Post to channel
│   └── teams_list_channel_messages — Read channel messages
│
└── Chats
    ├── teams_list_chats            — List 1:1 and group chats
    ├── teams_send_chat_message     — Send chat message
    ├── teams_list_chat_messages    — Read chat messages
    ├── teams_create_chat           — Create new chat
    └── teams_get_chat_members      — Get chat members
```

### 4.10 Google Cloud Storage (16 Tools)

Bucket and object management for Google Cloud Storage.

```
GCS Tools
├── Bucket Management
│   ├── gcs_list_buckets        — List all buckets
│   ├── gcs_get_bucket          — Get bucket details
│   ├── gcs_create_bucket       — Create bucket
│   └── gcs_delete_bucket       — Delete bucket
│
├── Object Operations
│   ├── gcs_list_objects         — List objects in bucket
│   ├── gcs_upload_object        — Upload object
│   ├── gcs_download_object      — Download object
│   ├── gcs_delete_object        — Delete object
│   ├── gcs_copy_object          — Copy object
│   ├── gcs_move_object          — Move object
│   ├── gcs_rename_object        — Rename object
│   └── gcs_get_object_metadata  — Get object metadata
│
├── Access Control
│   ├── gcs_make_object_public   — Make object public
│   └── gcs_make_object_private  — Make object private
│
├── Signed URLs
│   └── gcs_generate_signed_url  — Generate temporary signed URL
│
└── Bulk Operations
    └── gcs_batch_delete_objects  — Batch delete multiple objects
```

### 4.11 Meeting Transcription (3 Tools)

Search, view, and AI-summarize meeting transcription files.

```
Meeting Transcription Tools
├── list_meeting_transcriptions        — Search & list transcription files
├── open_meeting_transcription_file    — Read full transcription content
└── summarize_meeting_transcription    — AI-powered summary generation
```

---

## 5. Frontend Widgets & UI

### 5.1 Overall Layout

```
┌──────────────────────────────────────────────────────────────────┐
│                         HEADER BAR                               │
│  [Logo]  Multi-Service AI Agent    [Capabilities] [Settings]     │
├──────────────────────────────────────────────────────────────────┤
│                                                                  │
│  ┌────────────────────────────────┐  ┌────────────────────────┐  │
│  │                                │  │                        │  │
│  │       SERVICE PANELS           │  │      CHAT WINDOW       │  │
│  │                                │  │                        │  │
│  │  ┌──────────────────────────┐  │  │  ┌──────────────────┐  │  │
│  │  │ Gmail         [Connect] │  │  │  │  Message Feed    │  │  │
│  │  ├──────────────────────────┤  │  │  │                  │  │  │
│  │  │ Calendar      [Connect] │  │  │  │  User: "Send an  │  │  │
│  │  ├──────────────────────────┤  │  │  │  email to..."    │  │  │
│  │  │ Drive         [Connect] │  │  │  │                  │  │  │
│  │  ├──────────────────────────┤  │  │  │  Agent: "Done!   │  │  │
│  │  │ Sheets        [Connect] │  │  │  │  Email sent."    │  │  │
│  │  ├──────────────────────────┤  │  │  │                  │  │  │
│  │  │ Docs          [Connect] │  │  │  └──────────────────┘  │  │
│  │  ├──────────────────────────┤  │  │                        │  │
│  │  │ Google Chat   [Connect] │  │  │  ┌──────────────────┐  │  │
│  │  ├──────────────────────────┤  │  │  │  [Attach Files]  │  │  │
│  │  │ GitHub        [Connect] │  │  │  │  [Message Input]  │  │  │
│  │  ├──────────────────────────┤  │  │  │  [Send Button]   │  │  │
│  │  │ Outlook       [Connect] │  │  │  └──────────────────┘  │  │
│  │  ├──────────────────────────┤  │  │                        │  │
│  │  │ Teams         [Connect] │  │  └────────────────────────┘  │
│  │  ├──────────────────────────┤  │                              │
│  │  │ GCS           [Connect] │  │                              │
│  │  ├──────────────────────────┤  │                              │
│  │  │ Meeting Trans.[Connect] │  │                              │
│  │  ├──────────────────────────┤  │                              │
│  │  │ Timer Tasks   [Connect] │  │                              │
│  │  └──────────────────────────┘  │                              │
│  └────────────────────────────────┘                              │
└──────────────────────────────────────────────────────────────────┘
```

### 5.2 Widget Details

Each service panel has two states:

**Disconnected State:**
- Service name & icon
- "Connect" button to initiate OAuth
- Brief setup instructions (API keys, scopes needed)

**Connected State:**
- Green status indicator
- List of available tools for that service
- Quick-action buttons (service-specific)
- "Disconnect" option

### 5.3 Capabilities Modal

A comprehensive modal dialog showing all available tools grouped by service tabs:

```
┌─────────────────────────────────────────────────┐
│               CAPABILITIES                   [X] │
├─────────────────────────────────────────────────┤
│ [Gmail] [Calendar] [Drive] [Sheets] [Docs] ...  │
├─────────────────────────────────────────────────┤
│                                                  │
│  Gmail Tools (25 available)                      │
│  ──────────────────────────                      │
│  • send_email — Send emails with attachments     │
│  • search_emails — Search with query syntax      │
│  • read_email — Read full email content          │
│  • list_emails — List inbox messages             │
│  • ...                                           │
│                                                  │
└─────────────────────────────────────────────────┘
```

### 5.4 Chat Window Features

| Feature | Description |
|---|---|
| Message History | Up to 24 messages displayed |
| File Attachments | Up to 10 files per message |
| Service Filter | Filter chat by service context (all/gmail/calendar/etc.) |
| Streaming Responses | Real-time token-by-token display |
| Markdown Rendering | Rich text formatting in responses |
| Tool Activity | Shows which tools the agent is calling |

---

## 6. Meeting Transcription

### Overview

The meeting transcription feature provides the ability to search, browse, and AI-summarize meeting transcript files. The system processes transcription text through an intelligent chunked summarization pipeline to handle transcripts of any length.

### Architecture

```
┌──────────────────────────────────────────────────────────────────┐
│                  MEETING TRANSCRIPTION PIPELINE                  │
│                                                                  │
│  ┌──────────────┐    ┌───────────────┐    ┌──────────────────┐  │
│  │  Transcription│    │  Chunk        │    │  OpenAI GPT-4    │  │
│  │  Files        │───▶│  Splitter     │───▶│  Summarizer      │  │
│  │  (on disk)    │    │  (45K chars)  │    │                  │  │
│  └──────────────┘    └───────────────┘    └────────┬─────────┘  │
│                                                     │            │
│                                            Multiple chunk        │
│                                            summaries             │
│                                                     │            │
│                                                     ▼            │
│                      ┌───────────────────────────────────────┐   │
│                      │  Merge & Refine                       │   │
│                      │  (Final cohesive summary)             │   │
│                      │                                       │   │
│                      │  Outputs:                             │   │
│                      │  • Key Topics & Decisions             │   │
│                      │  • Action Items                       │   │
│                      │  • Important Highlights               │   │
│                      │  • Deep Links (Google Docs)           │   │
│                      └───────────────────────────────────────┘   │
└──────────────────────────────────────────────────────────────────┘
```

### Summarization Pipeline

1. **Input:** Raw transcription text (supports up to 500K characters)
2. **Short Path:** If text < 120K characters → single-pass summarization
3. **Long Path:** If text > 120K characters:
   - Split into 45K-character chunks
   - Each chunk summarized independently via `summarizeMeetingTranscriptChunkWithOpenAI()`
   - Chunk summaries merged and refined via `refineMergedMeetingSummaryWithOpenAI()`
4. **Output:** Structured summary with key topics, decisions, action items, and Google Docs deep links for important topics

### Tools

| Tool | Description |
|---|---|
| `list_meeting_transcriptions` | Search and list available transcription files (max 50 results) |
| `open_meeting_transcription_file` | Read the full content of a transcription file |
| `summarize_meeting_transcription` | Generate AI-powered summary with action items and key decisions |

---

## 7. AI Chat Engine

### Architecture

```
┌────────────────────────────────────────────────────────────────┐
│                      AI CHAT ENGINE                            │
│                                                                │
│  ┌──────────────┐                                              │
│  │ User Message  │                                             │
│  └──────┬───────┘                                              │
│         │                                                      │
│         ▼                                                      │
│  ┌──────────────────────────────────────────┐                  │
│  │         INTENT DETECTION                 │                  │
│  │                                          │                  │
│  │  • Meeting intent → Calendar tools       │                  │
│  │  • Document edit → Docs (not Sheets)     │                  │
│  │  • Email signature → User name hints     │                  │
│  │  • Self-only attendee filtering          │                  │
│  └──────────────┬───────────────────────────┘                  │
│                 │                                               │
│                 ▼                                               │
│  ┌──────────────────────────────────────────┐                  │
│  │         OPENAI TOOL CALLING              │                  │
│  │                                          │                  │
│  │  Model: gpt-4o-mini (primary)            │                  │
│  │  Fallback: gpt-4o                        │                  │
│  │  Temperature: 0.2                        │                  │
│  │  Max Output Tokens: 1500                 │                  │
│  │                                          │                  │
│  │  System Prompt + Available Tools         │                  │
│  │  → Model selects tools to call           │                  │
│  │  → Backend executes tools                │                  │
│  │  → Results fed back to model             │                  │
│  │  → Model generates final response        │                  │
│  └──────────────┬───────────────────────────┘                  │
│                 │                                               │
│                 ▼                                               │
│  ┌──────────────────────────────────────────┐                  │
│  │         RESPONSE DELIVERY                │                  │
│  │                                          │                  │
│  │  • Streaming (SSE) or JSON response      │                  │
│  │  • Chat history updated (max 14 msgs)    │                  │
│  │  • Tool results truncated if needed      │                  │
│  └──────────────────────────────────────────┘                  │
└────────────────────────────────────────────────────────────────┘
```

### Intelligent Routing Rules

The AI engine includes built-in routing intelligence:

| Rule | Behavior |
|---|---|
| Meeting Creation | Detects meeting intent and routes to `create_meet_event` with auto-generated Meet link |
| Document Editing | Prefers Google Docs over Sheets for document edit requests |
| Self-Only Attendee | When user says "schedule for me", only adds the authenticated user |
| Email Signatures | Injects user first name for email signature personalization |
| Suggestion Rules | Uses comprehensive linguistic cue analysis for smart suggestions |

### Chat History Management

| Parameter | Value |
|---|---|
| Max messages in context | 14 |
| Max chars per message | 1,800 |
| Max total history chars | 12,000 |
| Max history items stored | 40 |
| Max tool result chars | 25,000 |

### Retry & Fallback Logic

```
Attempt 1: gpt-4o-mini → Success ✓
Attempt 2: gpt-4o-mini (retry) → if rate limited
Attempt 3: gpt-4o (fallback) → if primary fails
Attempt 4: gpt-4o (retry) → final attempt
```

---

## 8. Authentication & OAuth

### OAuth Flow Diagram

```
┌────────┐     ┌────────────┐     ┌──────────────┐     ┌────────────┐
│  User  │────▶│  Frontend  │────▶│   Backend    │────▶│  Provider  │
│        │     │  Click     │     │  /auth       │     │  (Google/  │
│        │     │  Connect   │     │  endpoint    │     │  GitHub/   │
│        │     │            │     │              │     │  Microsoft)│
└────────┘     └────────────┘     └──────┬───────┘     └─────┬──────┘
                                         │                    │
                                         │   Auth URL +       │
                                         │   State Token      │
                                         │◀───────────────────│
                                         │                    │
                    Browser Redirect ────▶│                    │
                                         │  User Consents ───▶│
                                         │                    │
                                         │  Callback with     │
                                         │  Auth Code         │
                                         │◀───────────────────│
                                         │                    │
                                         │  Exchange for      │
                                         │  Access Token      │
                                         │───────────────────▶│
                                         │                    │
                                         │  Token Response    │
                                         │◀───────────────────│
                                         │                    │
                                  Store token to disk          │
                                  Initialize API client        │
                                         │                    │
                                         ▼                    │
                                  ┌──────────────┐            │
                                  │  Service      │            │
                                  │  Ready ✓      │            │
                                  └──────────────┘            │
```

### Provider Details

| Provider | Scopes | Token Storage | Refresh |
|---|---|---|---|
| **Google** | Gmail, Calendar, Chat, Drive, Sheets, Docs | `~/.gmail-mcp/token.json` | Auto-refresh on expiry |
| **GitHub** | repo, read:user, user:email, notifications, gist | `~/.gmail-mcp/github-token.json` | PAT-based (no expiry) |
| **Microsoft** | Mail, Teams, Chat, OneDrive, offline_access | `~/.gmail-mcp/outlook-token.json` | Auto-refresh on expiry |

### Security Measures

- **State parameter** with 10-minute TTL prevents CSRF attacks
- **Scope validation** per-client ensures minimum necessary permissions
- **Token encryption** at rest via filesystem permissions
- **Automatic token refresh** before expiry for seamless UX

---

## 9. API Endpoints

### Status Endpoints

| Method | Path | Description |
|---|---|---|
| GET | `/api/gmail/status` | Gmail connection status |
| GET | `/api/calendar/status` | Calendar connection status |
| GET | `/api/drive/status` | Drive connection status |
| GET | `/api/sheets/status` | Sheets connection status |
| GET | `/api/docs/status` | Docs connection status |
| GET | `/api/gchat/status` | Google Chat connection status |
| GET | `/api/github/status` | GitHub connection status |
| GET | `/api/outlook/status` | Outlook connection status |
| GET | `/api/teams/status` | Teams connection status |
| GET | `/api/gcs/status` | GCS connection status |
| GET | `/api/meeting-transcription/status` | Meeting transcription status |
| GET | `/api/sheets-mcp/status` | Sheets MCP bridge status |

### Authentication Endpoints

| Method | Path | Description |
|---|---|---|
| GET | `/api/gmail/auth` | Start Google OAuth flow |
| GET | `/api/calendar/connect` | Connect Calendar (shared Google OAuth) |
| GET | `/api/gchat/connect` | Connect Google Chat |
| GET | `/api/drive/connect` | Connect Google Drive |
| GET | `/api/sheets/connect` | Connect Google Sheets |
| GET | `/api/github/auth` | Start GitHub OAuth flow |
| POST | `/api/github/connect` | Connect with GitHub PAT |
| POST | `/api/github/disconnect` | Disconnect GitHub |
| GET | `/api/outlook/auth` | Start Microsoft OAuth flow |
| POST | `/api/outlook/disconnect` | Disconnect Outlook |
| GET | `/oauth2callback` | Google OAuth callback |
| GET | `/github/callback` | GitHub OAuth callback |
| GET | `/outlook/callback` | Microsoft OAuth callback |

### Chat & AI Endpoints

| Method | Path | Description |
|---|---|---|
| POST | `/api/chat` | Send message to AI agent (JSON response) |
| POST | `/api/chat/stream` | Send message with streaming (SSE response) |

### Tool Management

| Method | Path | Description |
|---|---|---|
| GET | `/api/tools` | List all available tools and their schemas |

### File Operations

| Method | Path | Description |
|---|---|---|
| POST | `/api/upload` | Upload file (multipart form) |
| GET | `/api/upload/:fileId` | Retrieve uploaded file metadata |
| DELETE | `/api/upload/:fileId` | Delete uploaded file |
| GET | `/api/download/:filename` | Download local file |
| GET | `/api/drive/download/:fileId` | Download from Google Drive |
| GET | `/api/gcs/download/:bucket/:objectName` | Download from GCS |

### Meeting Endpoints

| Method | Path | Description |
|---|---|---|
| GET | `/api/meet/metadata` | Get meeting metadata |
| POST | `/api/meet/finalize` | Process meeting transcript |
| POST | `/api/meet/share` | Share meeting notes |

### Timer / Scheduled Task Endpoints

| Method | Path | Description |
|---|---|---|
| GET | `/api/timer-tasks/status` | Timer service status |
| GET | `/api/timer-tasks` | List all scheduled tasks |
| POST | `/api/timer-tasks` | Create a scheduled task |
| PATCH | `/api/timer-tasks/:id` | Update a scheduled task |
| DELETE | `/api/timer-tasks/:id` | Delete a scheduled task |
| POST | `/api/timer-tasks/:id/run` | Manually run a task |

---

## 10. Security Architecture

### Defense Layers

```
┌─────────────────────────────────────────────────────────────┐
│  LAYER 1: Network Access Control                            │
│  • Local-only by default (127.0.0.1)                        │
│  • ALLOW_REMOTE_API flag for remote access                  │
│  • Loopback address validation                              │
├─────────────────────────────────────────────────────────────┤
│  LAYER 2: CORS Policy                                       │
│  • Restricted to localhost variants                         │
│  • Chrome extension origin matching                         │
│  • Preflight request handling                               │
├─────────────────────────────────────────────────────────────┤
│  LAYER 3: Rate Limiting                                     │
│  • API: 30 requests/minute                                  │
│  • Chat: 15 requests/minute                                 │
│  • Uploads: 10 requests/minute                              │
├─────────────────────────────────────────────────────────────┤
│  LAYER 4: Security Headers                                  │
│  • X-Content-Type-Options: nosniff                          │
│  • X-Frame-Options: DENY                                    │
│  • X-XSS-Protection: 1; mode=block                          │
│  • Referrer-Policy: strict-origin-when-cross-origin         │
│  • Content-Security-Policy                                  │
│  • Permissions-Policy: camera=(), geolocation=()            │
├─────────────────────────────────────────────────────────────┤
│  LAYER 5: Input Validation                                  │
│  • Path traversal prevention on file operations             │
│  • OAuth state parameter validation (10-min TTL)            │
│  • File size limits (5GB configurable)                      │
│  • Secure filename generation (UUID + timestamp)            │
├─────────────────────────────────────────────────────────────┤
│  LAYER 6: Data Protection                                   │
│  • Token storage in user home directory                     │
│  • Automatic file cleanup (15-min TTL)                      │
│  • No persistent database (reduced attack surface)          │
└─────────────────────────────────────────────────────────────┘
```

---

## 11. Configuration & Environment

### Environment Variables

```bash
# ─── AI Configuration ───────────────────────────────
OPENAI_API_KEY=                    # Required — OpenAI API key
OPENAI_MODEL=gpt-4o-mini           # Primary model
OPENAI_FALLBACK_MODEL=gpt-4o       # Fallback model
OPENAI_TEMPERATURE=0.2             # Response determinism (0-1)
OPENAI_MAX_OUTPUT_TOKENS=1500      # Max tokens per response
OPENAI_CHAT_MAX_RETRIES=2          # Retry attempts

# ─── Server Configuration ───────────────────────────
PORT=3000                           # Server port
HOST=127.0.0.1                      # Bind address
ALLOW_REMOTE_API=false              # Allow non-local access

# ─── Google OAuth ────────────────────────────────────
GOOGLE_CLIENT_ID=                   # Google Cloud Console
GOOGLE_CLIENT_SECRET=               # Google Cloud Console

# ─── GitHub OAuth ────────────────────────────────────
GITHUB_CLIENT_ID=                   # GitHub Developer Settings
GITHUB_CLIENT_SECRET=               # GitHub Developer Settings
GITHUB_REDIRECT_URI=http://localhost:3000/github/callback

# ─── Microsoft OAuth ────────────────────────────────
OUTLOOK_CLIENT_ID=                  # Azure AD App Registration
OUTLOOK_CLIENT_SECRET=              # Azure AD App Registration
OUTLOOK_REDIRECT_URI=http://localhost:3000/outlook/callback

# ─── MCP Bridge ─────────────────────────────────────
SHEETS_MCP_ENABLED=true             # Enable Sheets MCP tools
SHEETS_MCP_COMMAND=node             # MCP subprocess command
SHEETS_MCP_ARGS=                    # Path to MCP server
SHEETS_MCP_CREDS_DIR=               # MCP credentials directory
```

---

## 12. Data Flow & Tool Execution

### Complete Request Lifecycle

```
 Step 1: User Input
 ┌─────────────────┐
 │ "Send an email   │
 │  to john@..."    │
 └────────┬────────┘
          │
 Step 2: POST /api/chat
          │
          ▼
 ┌─────────────────────────────────────────────────────┐
 │ Backend builds OpenAI request:                      │
 │                                                     │
 │  {                                                  │
 │    model: "gpt-4o-mini",                            │
 │    messages: [system_prompt, ...history, user_msg],  │
 │    tools: [all_connected_tool_schemas],              │
 │    tool_choice: "auto"                              │
 │  }                                                  │
 └────────────────────┬────────────────────────────────┘
                      │
 Step 3: OpenAI selects tool
                      │
                      ▼
 ┌─────────────────────────────────────────────────────┐
 │ OpenAI Response:                                    │
 │                                                     │
 │  tool_calls: [{                                     │
 │    function: {                                      │
 │      name: "send_email",                            │
 │      arguments: {                                   │
 │        to: "john@example.com",                      │
 │        subject: "Hello",                            │
 │        body: "..."                                  │
 │      }                                              │
 │    }                                                │
 │  }]                                                 │
 └────────────────────┬────────────────────────────────┘
                      │
 Step 4: Execute tool
                      │
                      ▼
 ┌─────────────────────────────────────────────────────┐
 │ Backend executes send_email():                      │
 │  → Validates parameters                             │
 │  → Calls Gmail API                                  │
 │  → Returns result                                   │
 └────────────────────┬────────────────────────────────┘
                      │
 Step 5: Feed result back to OpenAI
                      │
                      ▼
 ┌─────────────────────────────────────────────────────┐
 │ OpenAI generates natural language response:         │
 │                                                     │
 │ "I've sent the email to john@example.com            │
 │  with subject 'Hello'. ✓"                           │
 └────────────────────┬────────────────────────────────┘
                      │
 Step 6: Return to frontend
                      │
                      ▼
 ┌─────────────────┐
 │ Chat message    │
 │ displayed to    │
 │ user            │
 └─────────────────┘
```

---

## 13. Scheduled Tasks & Timer

### Overview

The timer system allows users to schedule recurring AI agent tasks.

```
┌───────────────────────────────────────────────────┐
│                TIMER TASK SYSTEM                  │
│                                                   │
│  ┌─────────────────────────────────────────────┐  │
│  │  Task Definition                            │  │
│  │  {                                          │  │
│  │    id: "uuid",                              │  │
│  │    name: "Daily Standup Summary",           │  │
│  │    time: "09:00",                           │  │
│  │    instruction: "List my calendar events",  │  │
│  │    enabled: true,                           │  │
│  │    createdAt: "2026-02-20T..."              │  │
│  │  }                                          │  │
│  └─────────────────────────────────────────────┘  │
│                                                   │
│  Persistence: ~/.gmail-mcp/scheduled-tasks.json   │
│  Execution: Runs the instruction through /api/chat│
│  Controls: Enable/Disable, Manual Run, Delete     │
└───────────────────────────────────────────────────┘
```

---

## 14. File Management

### Upload & Download Pipeline

```
┌──────────┐     ┌──────────────┐     ┌───────────────┐
│  User    │────▶│  POST        │────▶│  Multer       │
│  Upload  │     │  /api/upload │     │  Middleware    │
└──────────┘     └──────────────┘     └───────┬───────┘
                                              │
                                              ▼
                                    ┌───────────────────┐
                                    │  File Storage     │
                                    │  ./uploads/       │
                                    │                   │
                                    │  Filename: UUID   │
                                    │  + timestamp      │
                                    │  + original ext   │
                                    │                   │
                                    │  Metadata stored  │
                                    │  in uploadedFiles │
                                    │  Map (in-memory)  │
                                    └───────┬───────────┘
                                            │
                                   ┌────────┴────────┐
                                   │                 │
                                   ▼                 ▼
                          ┌──────────────┐  ┌──────────────┐
                          │  Attach to   │  │  Auto-cleanup│
                          │  Chat Msg    │  │  after 15min │
                          └──────────────┘  └──────────────┘
```

### File Limits

| Parameter | Value |
|---|---|
| Max file size | 5 GB (configurable) |
| Max attached files per message | 10 |
| File TTL | 15 minutes |
| Cleanup interval | Every 10 minutes |
| Filename format | `{uuid}-{timestamp}.{ext}` |

---

## 15. Performance & Limits

### System Limits Reference

| Category | Parameter | Value |
|---|---|---|
| **Rate Limits** | API requests | 30/minute |
| | Chat requests | 15/minute |
| | Upload requests | 10/minute |
| **Chat** | Max messages in context | 14 |
| | Max chars per message | 1,800 |
| | Max total history chars | 12,000 |
| | Max stored history items | 40 |
| | Max frontend messages | 24 |
| **Tool Results** | Max result chars | 25,000 |
| | Max string value | 900 chars |
| | Max array items | 20 |
| | Max object keys | 50 |
| **Transcription** | Max text chars | 500,000 |
| | Chunk size | 45,000 chars |
| | Single-pass threshold | 120,000 chars |
| | Max search results | 50 |
| **Files** | Max upload size | 5 GB |
| | Max attachments/message | 10 |
| | File TTL | 15 minutes |
| **OpenAI** | Max output tokens | 1,500 |
| | Max retries | 4 (2 per model) |
| | Temperature | 0.2 |

---

## 16. Dependencies

### NPM Packages

| Package | Version | Purpose |
|---|---|---|
| `express` | ^4.18.2 | HTTP server framework |
| `openai` | ^4.28.0 | OpenAI API client |
| `googleapis` | ^129.0.0 | Google APIs (Gmail, Calendar, Drive, Sheets, Docs, Chat) |
| `octokit` | ^5.0.5 | GitHub REST API client |
| `@google-cloud/storage` | latest | Google Cloud Storage client |
| `@modelcontextprotocol/sdk` | ^1.26.0 | MCP client for tool bridge |
| `@isaacphi/mcp-gdrive` | ^0.2.0 | MCP server for Sheets/Drive |
| `ws` | ^8.18.0 | WebSocket server |
| `cors` | ^2.8.5 | CORS middleware |
| `multer` | ^2.0.2 | File upload handling |
| `express-rate-limit` | ^8.2.1 | Rate limiting |
| `dotenv` | ^16.3.1 | Environment variable loading |
| `uuid` | latest | Unique ID generation |

---

## Appendix A: File Structure

```
MCP-tools/
├── server.js                  # Main backend (10,134 lines)
├── package.json               # Node.js config & dependencies
├── .env                       # Environment variables (git-ignored)
├── .env.example               # Environment template
├── .gitignore                 # Git ignore rules
├── README.md                  # Project readme
├── uploads/                   # Temporary file storage
│
├── public/                    # Frontend application
│   ├── index.html             # Main UI layout (82.4 KB)
│   ├── app.js                 # Frontend logic (140.6 KB)
│   ├── styles.css             # Main styles (37.1 KB)
│   └── styles-modal.css       # Modal styles (9.7 KB)
│
└── ~/.gmail-mcp/              # User config directory
    ├── token.json             # Google OAuth tokens
    ├── github-token.json      # GitHub OAuth tokens
    ├── outlook-token.json     # Microsoft OAuth tokens
    └── scheduled-tasks.json   # Timer task persistence
```

---

## Appendix B: Service Connection State Diagram

```
                    ┌──────────┐
                    │          │
                    │  IDLE    │ (No token on disk)
                    │          │
                    └────┬─────┘
                         │
                    User clicks
                    "Connect"
                         │
                         ▼
                    ┌──────────┐
                    │          │
                    │  AUTH    │ (OAuth flow in progress)
                    │  PENDING │
                    │          │
                    └────┬─────┘
                         │
                    OAuth callback
                    received
                         │
                         ▼
                    ┌──────────┐
                    │          │
                    │ CONNECTED│ (API client initialized)
                    │    ✓     │ (Tools available)
                    │          │
                    └────┬─────┘
                         │
                    Token expired │ User disconnects
                         │                │
                         ▼                ▼
                    ┌──────────┐    ┌──────────┐
                    │ REFRESH  │    │  IDLE    │
                    │ (auto)   │    │          │
                    └────┬─────┘    └──────────┘
                         │
                    New token
                    obtained
                         │
                         ▼
                    ┌──────────┐
                    │ CONNECTED│
                    │    ✓     │
                    └──────────┘
```

---

*Document generated on February 20, 2026. This is a living document and should be updated as new services and features are added to the platform.*
