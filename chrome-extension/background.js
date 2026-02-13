/**
 * Google Meet Note-Taker Background Script
 * Manages WebSocket connection and session lifecycle
 */

console.log('[Background] Service worker started');

// Configuration
const WS_URL = 'ws://localhost:3000/meet-notes';
const BACKEND_URL = 'http://localhost:3000';
const RECONNECT_INTERVALS = [1000, 2000, 5000, 10000, 30000]; // Exponential backoff

// State
let ws = null;
let reconnectAttempts = 0;
let reconnectTimeout = null;
let currentSession = null;
let captionBuffer = []; // Buffer captions when disconnected
let isConnecting = false;

/**
 * Generate unique session ID
 */
function generateSessionId() {
  return `session_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
}

/**
 * Connect to WebSocket server
 */
function connectWebSocket() {
  if (ws && (ws.readyState === WebSocket.OPEN || ws.readyState === WebSocket.CONNECTING)) {
    console.log('[Background] WebSocket already connected or connecting');
    return;
  }

  if (isConnecting) {
    console.log('[Background] Connection attempt already in progress');
    return;
  }

  isConnecting = true;
  console.log('[Background] Connecting to WebSocket:', WS_URL);

  try {
    ws = new WebSocket(WS_URL);

    ws.onopen = () => {
      console.log('[Background] WebSocket connected');
      isConnecting = false;
      reconnectAttempts = 0;

      // Flush buffered captions
      if (captionBuffer.length > 0 && currentSession) {
        console.log(`[Background] Flushing ${captionBuffer.length} buffered captions`);
        captionBuffer.forEach(caption => {
          sendMessage({
            type: 'caption',
            sessionId: currentSession.sessionId,
            caption: caption
          });
        });
        captionBuffer = [];
      }

      // Notify popup
      broadcastToPopup({ type: 'WS_CONNECTED' });
    };

    ws.onmessage = (event) => {
      try {
        const data = JSON.parse(event.data);
        console.log('[Background] Message received:', data);

        // Broadcast to popup
        broadcastToPopup(data);

        // Handle specific message types
        if (data.type === 'session_created') {
          console.log('[Background] Session created:', data.sessionId);
        }
        else if (data.type === 'caption_added') {
          // Real-time caption update from server
          console.log('[Background] Caption added:', data.caption);
        }
        else if (data.type === 'error') {
          console.error('[Background] Server error:', data.message);
        }
      } catch (err) {
        console.error('[Background] Error parsing message:', err);
      }
    };

    ws.onerror = (error) => {
      console.error('[Background] WebSocket error:', error);
      isConnecting = false;
    };

    ws.onclose = () => {
      console.log('[Background] WebSocket disconnected');
      isConnecting = false;
      ws = null;

      // Notify popup
      broadcastToPopup({ type: 'WS_DISCONNECTED' });

      // Auto-reconnect if there's an active session
      if (currentSession) {
        scheduleReconnect();
      }
    };

  } catch (err) {
    console.error('[Background] Error creating WebSocket:', err);
    isConnecting = false;
    if (currentSession) {
      scheduleReconnect();
    }
  }
}

/**
 * Schedule reconnection with exponential backoff
 */
function scheduleReconnect() {
  if (reconnectTimeout) {
    clearTimeout(reconnectTimeout);
  }

  const delay = RECONNECT_INTERVALS[Math.min(reconnectAttempts, RECONNECT_INTERVALS.length - 1)];
  reconnectAttempts++;

  console.log(`[Background] Reconnecting in ${delay}ms (attempt ${reconnectAttempts})`);

  reconnectTimeout = setTimeout(() => {
    connectWebSocket();
  }, delay);
}

/**
 * Send message via WebSocket
 */
function sendMessage(data) {
  if (ws && ws.readyState === WebSocket.OPEN) {
    ws.send(JSON.stringify(data));
    console.log('[Background] Message sent:', data);
    return true;
  } else {
    console.warn('[Background] WebSocket not connected, cannot send message');
    return false;
  }
}

/**
 * Start capture session
 */
async function startSession(meetingMetadata) {
  console.log('[Background] Starting session:', meetingMetadata);

  const sessionId = generateSessionId();

  currentSession = {
    sessionId,
    metadata: meetingMetadata,
    startTime: new Date().toISOString(),
    captions: []
  };

  // Save to storage
  await chrome.storage.local.set({ currentSession });

  // Connect WebSocket if not connected
  if (!ws || ws.readyState !== WebSocket.OPEN) {
    connectWebSocket();
    // Wait for connection
    await new Promise(resolve => {
      const checkInterval = setInterval(() => {
        if (ws && ws.readyState === WebSocket.OPEN) {
          clearInterval(checkInterval);
          resolve();
        }
      }, 100);

      // Timeout after 5 seconds
      setTimeout(() => {
        clearInterval(checkInterval);
        resolve();
      }, 5000);
    });
  }

  // Send session start message
  sendMessage({
    type: 'session_start',
    sessionId,
    metadata: meetingMetadata
  });

  return sessionId;
}

/**
 * Add caption to current session
 */
async function addCaption(caption) {
  if (!currentSession) {
    console.warn('[Background] No active session, ignoring caption');
    return;
  }

  // Add to local storage
  currentSession.captions.push(caption);
  await chrome.storage.local.set({ currentSession });

  // Send to server via WebSocket
  const sent = sendMessage({
    type: 'caption',
    sessionId: currentSession.sessionId,
    caption
  });

  // Buffer if send failed
  if (!sent) {
    captionBuffer.push(caption);
    console.log('[Background] Caption buffered (disconnected)');
  }
}

/**
 * End capture session
 */
async function endSession() {
  if (!currentSession) {
    console.warn('[Background] No active session to end');
    return null;
  }

  console.log('[Background] Ending session:', currentSession.sessionId);

  // Send session end message
  sendMessage({
    type: 'session_end',
    sessionId: currentSession.sessionId
  });

  const sessionId = currentSession.sessionId;

  // Clear session
  currentSession = null;
  captionBuffer = [];
  await chrome.storage.local.remove('currentSession');

  // Close WebSocket
  if (ws) {
    ws.close();
    ws = null;
  }

  return sessionId;
}

/**
 * Broadcast message to all popup instances
 */
function broadcastToPopup(message) {
  chrome.runtime.sendMessage(message).catch(() => {
    // Popup might not be open, ignore error
  });
}

/**
 * Fetch meeting metadata from backend
 */
async function fetchMeetingMetadata(meetingCode) {
  try {
    const response = await fetch(`${BACKEND_URL}/api/meet/metadata?code=${meetingCode}`);
    const data = await response.json();

    if (data.error) {
      throw new Error(data.error);
    }

    return data.meeting;
  } catch (err) {
    console.error('[Background] Error fetching meeting metadata:', err);
    throw err;
  }
}

/**
 * Finalize meeting (generate summary and create doc)
 */
async function finalizeMeeting(sessionId) {
  try {
    const response = await fetch(`${BACKEND_URL}/api/meet/finalize`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({ sessionId })
    });

    const data = await response.json();

    if (data.error) {
      throw new Error(data.error);
    }

    return data;
  } catch (err) {
    console.error('[Background] Error finalizing meeting:', err);
    throw err;
  }
}

/**
 * Message handler
 */
chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  console.log('[Background] Message received:', message);

  // Handle caption captured from content script
  if (message.type === 'CAPTION_CAPTURED') {
    addCaption(message.caption)
      .then(() => {
        sendResponse({ success: true });
      })
      .catch(err => {
        console.error('[Background] Error adding caption:', err);
        sendResponse({ success: false, error: err.message });
      });
    return true; // Async response
  }

  // Handle session start request from popup
  else if (message.type === 'START_SESSION') {
    startSession(message.metadata)
      .then(sessionId => {
        sendResponse({ success: true, sessionId });

        // Tell content script to start capturing
        chrome.tabs.query({ active: true, currentWindow: true }, (tabs) => {
          if (tabs[0]) {
            chrome.tabs.sendMessage(tabs[0].id, { type: 'START_CAPTURE' });
          }
        });
      })
      .catch(err => {
        console.error('[Background] Error starting session:', err);
        sendResponse({ success: false, error: err.message });
      });
    return true; // Async response
  }

  // Handle session end request from popup
  else if (message.type === 'END_SESSION') {
    endSession()
      .then(sessionId => {
        // Tell content script to stop capturing
        chrome.tabs.query({ active: true, currentWindow: true }, (tabs) => {
          if (tabs[0]) {
            chrome.tabs.sendMessage(tabs[0].id, { type: 'STOP_CAPTURE' });
          }
        });

        sendResponse({ success: true, sessionId });
      })
      .catch(err => {
        console.error('[Background] Error ending session:', err);
        sendResponse({ success: false, error: err.message });
      });
    return true; // Async response
  }

  // Handle fetch metadata request from popup
  else if (message.type === 'FETCH_METADATA') {
    fetchMeetingMetadata(message.meetingCode)
      .then(metadata => {
        sendResponse({ success: true, metadata });
      })
      .catch(err => {
        sendResponse({ success: false, error: err.message });
      });
    return true; // Async response
  }

  // Handle finalize request from popup
  else if (message.type === 'FINALIZE_MEETING') {
    finalizeMeeting(message.sessionId)
      .then(result => {
        sendResponse({ success: true, result });
      })
      .catch(err => {
        sendResponse({ success: false, error: err.message });
      });
    return true; // Async response
  }

  // Handle get session request from popup
  else if (message.type === 'GET_SESSION') {
    sendResponse({ session: currentSession });
    return false;
  }

  // Handle meeting detected from content script
  else if (message.type === 'MEETING_DETECTED') {
    console.log('[Background] Meeting detected:', message.meetingCode, message.meetingTitle);
    // Store for later use
    chrome.storage.local.set({
      lastMeetingCode: message.meetingCode,
      lastMeetingTitle: message.meetingTitle
    });
    return false;
  }

  return false;
});

// Restore session on startup (in case of browser restart)
chrome.storage.local.get(['currentSession'], (result) => {
  if (result.currentSession) {
    console.log('[Background] Restoring session from storage');
    currentSession = result.currentSession;
    connectWebSocket();
  }
});

console.log('[Background] Service worker initialized');
