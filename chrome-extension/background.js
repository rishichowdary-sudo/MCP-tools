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
let pendingSessionStart = false;
let activeMeetTabId = null;

function delay(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

async function getMeetTabs() {
  const tabs = await chrome.tabs.query({ url: ['https://meet.google.com/*'] });
  return (tabs || []).filter(tab => Number.isInteger(tab.id));
}

async function sendMessageToMeetTab(tabId, message) {
  try {
    const response = await chrome.tabs.sendMessage(tabId, message);
    return { success: true, response };
  } catch (error) {
    return { success: false, error };
  }
}

async function sendCaptureCommandToMeetTabs(commandType, { retries = 5, retryDelayMs = 1200 } = {}) {
  let lastError = 'No Google Meet tab found';

  for (let attempt = 1; attempt <= retries; attempt += 1) {
    const tabs = await getMeetTabs();
    if (!tabs.length) {
      lastError = 'No Google Meet tab found';
      if (attempt < retries) await delay(retryDelayMs);
      continue;
    }

    const orderedTabs = [...tabs].sort((a, b) => {
      if (a.id === activeMeetTabId) return -1;
      if (b.id === activeMeetTabId) return 1;
      if (a.active && !b.active) return -1;
      if (!a.active && b.active) return 1;
      return 0;
    });

    for (const tab of orderedTabs) {
      const result = await sendMessageToMeetTab(tab.id, { type: commandType });
      if (result.success && result.response?.success !== false) {
        activeMeetTabId = tab.id;
        return tab.id;
      }
      lastError = result.error?.message || `Failed to send ${commandType} to tab ${tab.id}`;
    }

    if (attempt < retries) await delay(retryDelayMs);
  }

  throw new Error(lastError);
}

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

      // Ensure server receives session_start before any buffered captions.
      if (currentSession && pendingSessionStart) {
        const started = sendMessage({
          type: 'session_start',
          sessionId: currentSession.sessionId,
          metadata: currentSession.metadata
        });
        if (started) {
          pendingSessionStart = false;
        }
      }

      // Flush buffered captions
      if (captionBuffer.length > 0 && currentSession && !pendingSessionStart) {
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
        // Rebind session on reconnect so the server accepts future captions.
        pendingSessionStart = true;
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
  pendingSessionStart = true;

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

  // Send session start only if it wasn't already sent in ws.onopen.
  if (pendingSessionStart) {
    const started = sendMessage({
      type: 'session_start',
      sessionId,
      metadata: meetingMetadata
    });
    if (started) {
      pendingSessionStart = false;
    }
  }

  return sessionId;
}

/**
 * Add caption to current session
 */
async function addCaption(caption) {
  if (!currentSession) {
    // Service worker can restart; recover active session from storage.
    const restored = await chrome.storage.local.get(['currentSession']);
    if (restored?.currentSession) {
      currentSession = restored.currentSession;
      pendingSessionStart = true;
      if (!ws || ws.readyState !== WebSocket.OPEN) {
        connectWebSocket();
      }
    } else {
      console.warn('[Background] No active session, ignoring caption');
      return;
    }
  }

  // Send first so storage failures never block live forwarding.
  const sent = sendMessage({
    type: 'caption',
    sessionId: currentSession.sessionId,
    caption
  });

  // Buffer if send failed
  if (!sent) {
    captionBuffer.push(caption);
    console.log('[Background] Caption buffered (disconnected)');
    if (!ws || ws.readyState !== WebSocket.OPEN) {
      connectWebSocket();
    }
  }

  // Persist locally as best effort.
  currentSession.captions.push(caption);
  try {
    await chrome.storage.local.set({ currentSession });
  } catch (storageError) {
    console.warn('[Background] Failed to persist caption in storage:', storageError?.message || storageError);
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
  pendingSessionStart = false;
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
    const data = await parseBackendJsonResponse(response, 'fetch meeting metadata');

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

    const data = await parseBackendJsonResponse(response, 'finalize meeting');

    if (data.error) {
      throw new Error(data.error);
    }

    return data;
  } catch (err) {
    console.error('[Background] Error finalizing meeting:', err);
    throw err;
  }
}

async function parseBackendJsonResponse(response, operation) {
  const text = await response.text();
  let data;
  try {
    data = JSON.parse(text);
  } catch {
    const snippet = String(text || '').slice(0, 180).replace(/\s+/g, ' ').trim();
    throw new Error(`${operation} failed: backend returned non-JSON (${response.status})${snippet ? `: ${snippet}` : ''}`);
  }

  if (!response.ok) {
    throw new Error(data?.error || `${operation} failed with status ${response.status}`);
  }

  return data;
}

/**
 * Message handler
 */
chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  console.log('[Background] Message received:', message);

  // Handle caption captured from content script
  if (message.type === 'CAPTION_CAPTURED') {
    if (sender?.tab?.id) {
      activeMeetTabId = sender.tab.id;
    }
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
      .then(async (sessionId) => {
        try {
          await sendCaptureCommandToMeetTabs('START_CAPTURE', { retries: 6, retryDelayMs: 1000 });
        } catch (captureStartError) {
          console.warn('[Background] START_CAPTURE did not reach Meet tab:', captureStartError.message);
        }

        sendResponse({ success: true, sessionId });
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
      .then(async (sessionId) => {
        try {
          await sendCaptureCommandToMeetTabs('STOP_CAPTURE', { retries: 2, retryDelayMs: 500 });
        } catch (captureStopError) {
          console.warn('[Background] STOP_CAPTURE could not be confirmed:', captureStopError.message);
        }

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
    if (sender?.tab?.id) {
      activeMeetTabId = sender.tab.id;
    }
    // Store for later use
    chrome.storage.local.set({
      lastMeetingCode: message.meetingCode,
      lastMeetingTitle: message.meetingTitle
    });
    return false;
  }

  else if (message.type === 'CAPTURE_STARTED' || message.type === 'CAPTURE_STOPPED') {
    if (sender?.tab?.id) {
      activeMeetTabId = sender.tab.id;
    }
    return false;
  }

  return false;
});

// Restore session on startup (in case of browser restart)
chrome.storage.local.get(['currentSession'], (result) => {
  if (result.currentSession) {
    console.log('[Background] Restoring session from storage');
    currentSession = result.currentSession;
    pendingSessionStart = true;
    connectWebSocket();
  }
});

console.log('[Background] Service worker initialized');
