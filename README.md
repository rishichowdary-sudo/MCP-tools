# Google Meet Caption Note-Taker

This project includes:
- A Node.js backend (`server.js`)
- A Chrome extension (`chrome-extension/`) for capturing Google Meet captions

## Run The App

1. Open a terminal in the project root.
2. Install dependencies:

```bash
npm install
```

3. Start the backend server:

```bash
npm start
```

4. Confirm the server is running on `http://localhost:3000`.

## Add Extension In Chrome

1. Open Chrome and go to `chrome://extensions`.
2. Turn on **Developer mode** (top-right).
3. Click **Load unpacked**.
4. Select the `chrome-extension` folder from this project.
5. Pin the extension from the Chrome toolbar (optional, but recommended).

## Important Note

The extension is configured to connect to:
- Backend API: `http://localhost:3000`
- WebSocket: `ws://localhost:3000/meet-notes`

So keep `npm start` running while using the extension.
