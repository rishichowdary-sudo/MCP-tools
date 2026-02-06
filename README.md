# MCP Tools Agent UI

Web app that connects OpenAI chat with Gmail, Google Calendar, and GitHub tools.

## Prerequisites

- Node.js `18+`
- npm
- OpenAI API key
- Google OAuth client (for Gmail/Calendar features)
- GitHub Personal Access Token (optional, for GitHub features)

## 1) Install dependencies

```bash
npm install
```

## 2) Create `.env`

Create a `.env` file in the project root:

```env
OPENAI_API_KEY=your_openai_api_key_here
PORT=3000

# Google OAuth (required for Gmail/Calendar)
# You can use either GOOGLE_* or GMAIL_* names.
GOOGLE_CLIENT_ID=your_google_client_id
GOOGLE_CLIENT_SECRET=your_google_client_secret
```

## 3) Configure Google OAuth

In Google Cloud Console:

1. Enable Gmail API and Google Calendar API.
2. Create OAuth client credentials (Web application).
3. Add this redirect URI:
   `http://localhost:3000/oauth2callback`
   If you changed `PORT`, use that port in the URI.

## 4) Run the app

```bash
npm start
```

Open:

`http://localhost:3000`

## 5) Connect services in the UI

1. Click `Sign in with Google` to connect Gmail.
2. For Calendar, use the Calendar panel `Sign in with Google` if it asks again.
3. For GitHub, paste a Personal Access Token in the GitHub panel and connect.

## Notes

- Google tokens are saved locally at:
  `~/.gmail-mcp/token.json`
- GitHub token is saved locally at:
  `~/.gmail-mcp/github-token.json`
- If Calendar shows `Insufficient Permission`, reconnect from the Calendar panel to grant Calendar scope.
