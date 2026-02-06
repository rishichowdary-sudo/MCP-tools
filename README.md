# MCP Tools Agent UI

Web app that connects OpenAI chat with Gmail, Google Calendar, and GitHub tools.

## Prerequisites

- Node.js `18+`
- npm
- OpenAI API key
- Google OAuth client (for Gmail/Calendar features)
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

# Google OAuth (required for Gmail/Calendar)
# You can use either GOOGLE_* or GMAIL_* names.
GOOGLE_CLIENT_ID=your_google_client_id
GOOGLE_CLIENT_SECRET=your_google_client_secret

# GitHub OAuth (required for GitHub Sign in)
GITHUB_CLIENT_ID=your_github_client_id
GITHUB_CLIENT_SECRET=your_github_client_secret
# Optional override (default: http://localhost:3000/github/callback)
# GITHUB_REDIRECT_URI=http://localhost:3000/github/callback
```

## 3) Configure Google OAuth

In Google Cloud Console:

1. Enable Gmail API and Google Calendar API.
2. Create OAuth client credentials (Web application).
3. Add this redirect URI:
   `http://localhost:3000/oauth2callback`
   If you changed `PORT`, use that port in the URI.

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
2. For Calendar, use the Calendar panel `Sign in with Google` if it asks again.
3. Click `Sign in with GitHub` in the GitHub panel.
4. Use each panel's `Reauthenticate` button any time you want to switch accounts.

## Notes

- Google tokens are saved locally at:
  `~/.gmail-mcp/token.json`
- GitHub token is saved locally at:
  `~/.gmail-mcp/github-token.json`
- If Calendar shows `Insufficient Permission`, reconnect from the Calendar panel to grant Calendar scope.
