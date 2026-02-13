# Google Meet Note-Taker Chrome Extension

AI-powered meeting notes for Google Meet. Capture live captions, generate summaries, and automatically create shared Google Docs.

## Features

- ✅ **Real-time Caption Capture**: Automatically extracts Google Meet's live captions
- ✅ **AI Summaries**: GPT-4 generates key points, decisions, and action items
- ✅ **Auto-Documentation**: Creates Google Docs with full transcript + summary
- ✅ **Auto-Sharing**: Shares notes with meeting attendees from Calendar
- ✅ **Live Transcript View**: See captions in real-time in the extension popup
- ✅ **Offline Buffering**: Captions buffer locally if backend disconnects

## Architecture

```
Google Meet → Extension Content Script → WebSocket → Backend Server → OpenAI + Google Docs
```

## Setup

### 1. Prerequisites

- **Chrome Browser** (or any Chromium-based browser)
- **Backend Server** running at `http://localhost:3000` (see main README)
- **Google OAuth** configured with Calendar + Docs scopes
- **OpenAI API Key** for summary generation

### 2. Add Extension Icons

Before loading the extension, add icon files to `icons/` directory:

- `icon16.png` (16x16 pixels)
- `icon48.png` (48x48 pixels)
- `icon128.png` (128x128 pixels)

You can:
- Use an emoji-to-PNG converter (📝 notepad emoji)
- Download from [Flaticon](https://www.flaticon.com/) or [Icons8](https://icons8.com/)
- Create custom icons with Figma/Canva

Recommended colors: Blue (#4285F4) or Green (#34A853) to match Google Meet branding.

### 3. Load Extension in Chrome

1. Open Chrome and navigate to `chrome://extensions`
2. Enable **Developer mode** (toggle in top-right corner)
3. Click **Load unpacked**
4. Select the `chrome-extension/` folder
5. The extension icon should appear in your toolbar

### 4. Verify Backend Connection

1. Start the backend server:
   ```bash
   npm start
   ```

2. Ensure you see this log:
   ```
   WebSocket server initialized on path: /meet-notes
   ```

3. Click the extension icon — the footer should show:
   ```
   Backend: Connected ✓
   ```

## Usage

### Step 1: Join a Google Meet

1. Join or start a Google Meet call
2. **Enable captions**: Click the "CC" button in the bottom toolbar
3. Captions should appear at the bottom of the screen

### Step 2: Start Capture

1. Click the extension icon in the toolbar
2. The popup will fetch meeting details from your Calendar
3. Click **Start Capture**
4. Speak in the meeting — captions will appear in the extension popup

### Step 3: Stop & Save

1. When the meeting ends, click **Stop & Save**
2. The extension will:
   - Send captions to backend
   - Generate AI summary (3-5 seconds)
   - Create Google Doc with notes
   - Share doc with meeting attendees
3. Click **Open Google Doc** to view the notes

## How It Works

### Content Script (`meet-capture.js`)

- Injects into `meet.google.com/*` pages
- Uses `MutationObserver` to watch for caption DOM changes
- Extracts speaker name + caption text
- Sends to background script via `chrome.runtime.sendMessage`

### Background Script (`background.js`)

- Maintains WebSocket connection to `ws://localhost:3000/meet-notes`
- Buffers captions if disconnected, auto-reconnects with exponential backoff
- Manages session lifecycle (start → capture → end)
- Stores active session in Chrome Storage API

### Popup UI (`popup.html`)

- Start/Stop buttons for capture control
- Live transcript view (scrollable, auto-updates)
- Meeting metadata display (title, time, attendees)
- Success screen with Google Doc link

### Backend Integration

**WebSocket Server** (`server.js`):
- Path: `ws://localhost:3000/meet-notes`
- Messages: `session_start`, `caption`, `session_end`
- Stores sessions in-memory Map

**REST Endpoints**:
- `GET /api/meet/metadata?code=xyz` — Fetch meeting from Calendar
- `POST /api/meet/finalize` — Generate summary & create Google Doc

## Troubleshooting

### "Caption container not found"

**Fix**: Enable captions in Google Meet (CC button in toolbar)

### "Backend: Disconnected"

**Fix**: Ensure backend server is running at `http://localhost:3000`

```bash
npm start
```

### "Meeting not found in calendar"

**Fix**: Either:
- Create a Calendar event with the Meet link first
- Extension will use generic "Google Meet" title

### No captions appearing in popup

**Checks**:
1. Captions enabled in Meet? (CC button on)
2. Extension started? (Click "Start Capture")
3. Console errors? (Right-click extension icon → Inspect popup → Console tab)

### Captions not saving to Google Doc

**Checks**:
1. Google Docs connected? Check backend logs for "Google Docs: Connected"
2. Attendee emails correct? Check Calendar event attendees
3. OpenAI API key valid? Check `.env` file

## Development

### File Structure

```
chrome-extension/
├── manifest.json                    # Manifest V3 config
├── background.js                    # WebSocket client, session mgmt
├── content-scripts/
│   └── meet-capture.js             # Caption extraction
├── popup/
│   ├── popup.html                  # UI markup
│   ├── popup.js                    # UI logic
│   └── popup.css                   # Styling
└── icons/
    ├── icon16.png
    ├── icon48.png
    └── icon128.png
```

### Debugging

**Content Script**:
```
Open Meet tab → F12 → Console tab → Filter: "[Meet Capture]"
```

**Background Script**:
```
chrome://extensions → Extension details → Service Worker → Inspect
```

**Popup**:
```
Right-click extension icon → Inspect popup
```

### Testing Locally

1. Load extension in Chrome (Developer mode)
2. Join a test Meet: [meet.google.com/new](https://meet.google.com/new)
3. Enable captions (CC button)
4. Click extension → Start Capture
5. Speak or use auto-captions
6. Verify captions appear in popup
7. Click Stop & Save
8. Check Google Docs for created note

## Permissions

The extension requests these permissions in `manifest.json`:

- `activeTab` — Access current tab URL (to extract meeting code)
- `storage` — Store active session across browser restarts
- `tabs` — Query active tabs for content script injection
- `https://meet.google.com/*` — Access Meet pages for caption capture
- `http://localhost:3000/*` — Connect to backend WebSocket server

## Privacy

- **No data leaves your machine** except via your own backend server
- Captions are buffered locally if backend disconnects
- Extension does NOT upload recordings to third parties
- You control when capture starts/stops (manual process)

## Publishing (Optional)

To publish to Chrome Web Store:

1. Create a ZIP of the extension:
   ```bash
   cd chrome-extension
   zip -r meet-note-taker.zip * -x "*.md"
   ```

2. Visit [Chrome Web Store Developer Dashboard](https://chrome.google.com/webstore/devconsole)
3. Pay one-time $5 registration fee
4. Upload ZIP and submit for review

**Note**: If hosting backend on cloud server, update `WS_URL` and `BACKEND_URL` in `background.js`

## License

MIT License - See main project README

## Credits

Inspired by open source projects:
- [Meet-Script](https://github.com/RutvijDv/Meet-Script)
- [TranscripTonic](https://github.com/vivek-nexus/transcriptonic)
- [Recall.ai POC](https://github.com/recallai/google-meet-meeting-bot)
