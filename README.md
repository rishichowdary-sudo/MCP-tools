# MCP Tools Agent UI

Web app that connects OpenAI chat with Gmail, Google Calendar (including Google Meet links), Google Chat, Google Drive, Google Sheets, and GitHub tools.

## Prerequisites

- Node.js `18+`
- npm
- OpenAI API key
- Google OAuth client (for Gmail/Calendar/Chat/Drive/Sheets features)
- GitHub OAuth app (for GitHub features)

## 1) Install dependencies

```bash
npm install
```

## 2) Create `.env`

Create a `.env` file in the project root:

```env
OPENAI_API_KEY=your_openai_api_key_here
PORT=3000

# Google OAuth (required for Gmail/Calendar/Chat/Drive/Sheets)
# You can use either GOOGLE_* or GMAIL_* names.
GOOGLE_CLIENT_ID=your_google_client_id
GOOGLE_CLIENT_SECRET=your_google_client_secret

# Optional prebuilt Sheets MCP bridge (enabled by default)
# SHEETS_MCP_ENABLED=true
# SHEETS_MCP_COMMAND=node
# SHEETS_MCP_ARGS=node_modules/@isaacphi/mcp-gdrive/dist/index.js
# SHEETS_MCP_CREDS_DIR=C:\Users\<you>\.gmail-mcp\sheets-mcp

# GitHub OAuth (required for GitHub Sign in)
GITHUB_CLIENT_ID=your_github_client_id
GITHUB_CLIENT_SECRET=your_github_client_secret
# Optional override (default: http://localhost:3000/github/callback)
# GITHUB_REDIRECT_URI=http://localhost:3000/github/callback
```

## 3) Configure Google OAuth

In Google Cloud Console:

1. Enable Gmail API, Google Calendar API, Google Chat API, Google Drive API, and Google Sheets API.
2. Create OAuth client credentials (Web application).
3. Add this redirect URI:
   `http://localhost:3000/oauth2callback`
   If you changed `PORT`, use that port in the URI.

## 3b) Optional: Sheets MCP bridge

This project now includes a prebuilt MCP server package: `@isaacphi/mcp-gdrive`.

- It starts automatically on server boot.
- It exposes MCP tool names prefixed with `sheets_mcp__`.
- It stores MCP auth files in `SHEETS_MCP_CREDS_DIR` (default `~/.gmail-mcp/sheets-mcp`).

If you do not want this bridge, set:

```env
SHEETS_MCP_ENABLED=false
```

## 4) Configure GitHub OAuth

In GitHub Developer Settings:

1. Create a new OAuth App.
2. Set Homepage URL to:
   `http://localhost:3000`
3. Set Authorization callback URL to:
   `http://localhost:3000/github/callback`
   If you changed `PORT`, use that port in the callback.
4. Copy Client ID and Client Secret into `.env`.

## 5) Run the app

```bash
npm start
```

Open:

`http://localhost:3000`

## 6) Connect services in the UI

1. Click `Sign in with Google` to connect Gmail.
2. For Calendar, Google Chat, Google Drive, and Google Sheets, use each panel's `Sign in with Google` button if re-consent is needed.
3. Click `Sign in with GitHub` in the GitHub panel.
4. Use each panel's `Reauthenticate` button any time you want to switch accounts.

## Meet + Chat + Drive + Sheets Tools

- Calendar now supports:
  - `create_meet_event` to create events with Meet links
  - `add_meet_link_to_event` to add Meet links to existing events
  - `create_event` with `createMeetLink=true`
- Google Chat now supports:
  - `list_chat_spaces`
  - `send_chat_message`
  - `list_chat_messages`
- Google Drive now supports:
  - listing, creating, updating, moving, sharing, and downloading files
- Google Sheets now supports:
  - listing spreadsheets/tabs, reading/writing ranges, appending rows, and tab management
  - `update_timesheet_hours` for reliable date-based row updates (billing hours, task details, project/module fields)
  - additional MCP fallback tools prefixed with `sheets_mcp__`
- Timer Tasks widget now supports:
  - creating daily scheduled instructions (for any workflow, not just timesheets)
  - auto-running at configured `HH:MM` local time
  - run now, enable/disable, and delete actions from the UI

## Agentic Prompting (One Command, Multiple Tools)

Use a single goal-oriented command that clearly states:

1. End goal
2. Scope/time range
3. Constraints (for example: "do not delete")
4. Desired output format

Examples:

- `Find my 10 newest unread Gmail + Outlook emails, summarize each in 1 line, and draft one follow-up email for urgent items.`
- `Check today's calendar, find open PRs assigned to me, and schedule 30-minute review blocks this week.`
- `Create a Google Doc called Weekly Update, pull top action items from unread emails, and append them into the doc.`
- `List my spreadsheets updated this week, summarize tab names, and post the summary to my first Google Chat space.`

Tip:

- Ask for full execution in one message: `Do all required tool calls automatically and return a final summary with what succeeded, what failed, and next actions.`

## Notes

- Google tokens are saved locally at:
  `~/.gmail-mcp/token.json`
- GitHub token is saved locally at:
  `~/.gmail-mcp/github-token.json`
- Sheets MCP credentials are saved locally at:
  `~/.gmail-mcp/sheets-mcp`
- If Calendar shows `Insufficient Permission`, reconnect from the Calendar panel to grant Calendar scope.
