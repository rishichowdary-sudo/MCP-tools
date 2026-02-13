/**
 * Google Meet Note-Taker Popup Script
 * UI logic for the extension popup
 */

console.log('[Popup] Script loaded');

// DOM Elements
const startBtn = document.getElementById('startBtn');
const stopBtn = document.getElementById('stopBtn');
const statusIndicator = document.getElementById('statusIndicator');
const meetingInfo = document.getElementById('meetingInfo');
const meetingTitle = document.getElementById('meetingTitle');
const meetingTime = document.getElementById('meetingTime');
const meetingAttendees = document.getElementById('meetingAttendees');
const transcriptContainer = document.getElementById('transcriptContainer');
const transcriptContent = document.getElementById('transcriptContent');
const captionCount = document.getElementById('captionCount');
const successMessage = document.getElementById('successMessage');
const errorMessage = document.getElementById('errorMessage');
const errorText = document.getElementById('errorText');
const dismissError = document.getElementById('dismissError');
const loading = document.getElementById('loading');
const loadingText = document.getElementById('loadingText');
const successText = document.getElementById('successText');
const docLink = document.getElementById('docLink');
const connectionStatus = document.getElementById('connectionStatus');
const wsStatus = document.getElementById('wsStatus');

// State
let currentSession = null;
let captions = [];
let meetingMetadata = null;

/**
 * Initialize popup
 */
async function initialize() {
  console.log('[Popup] Initializing...');

  // Check for active session
  const response = await sendMessageToBackground({ type: 'GET_SESSION' });

  if (response.session) {
    console.log('[Popup] Active session found:', response.session);
    currentSession = response.session;
    captions = response.session.captions || [];
    meetingMetadata = response.session.metadata;

    showCapturingState();
  } else {
    // Try to fetch meeting metadata
    await fetchMeetingInfo();
  }

  // Set up event listeners
  setupEventListeners();
}

/**
 * Fetch meeting info from current tab
 */
async function fetchMeetingInfo() {
  try {
    const tabs = await chrome.tabs.query({ active: true, currentWindow: true });
    const currentTab = tabs[0];

    if (!currentTab.url.includes('meet.google.com')) {
      showError('Please open a Google Meet tab');
      return;
    }

    // Extract meeting code from URL
    const match = currentTab.url.match(/meet\.google\.com\/([a-z-]+)/);
    if (!match) {
      showError('Could not extract meeting code from URL');
      return;
    }

    const meetingCode = match[1];
    console.log('[Popup] Meeting code:', meetingCode);

    // Fetch metadata from backend
    showLoading('Fetching meeting info...');
    const response = await sendMessageToBackground({
      type: 'FETCH_METADATA',
      meetingCode
    });

    hideLoading();

    if (response.success && response.metadata) {
      meetingMetadata = response.metadata;
      displayMeetingInfo(response.metadata);
    } else {
      // No metadata found (meeting not in calendar)
      meetingMetadata = {
        title: 'Google Meet',
        meetingCode,
        startTime: new Date().toISOString(),
        attendees: []
      };
      displayMeetingInfo(meetingMetadata);
    }

  } catch (err) {
    hideLoading();
    console.error('[Popup] Error fetching meeting info:', err);
    showError('Failed to load meeting info: ' + err.message);
  }
}

/**
 * Display meeting information
 */
function displayMeetingInfo(metadata) {
  meetingInfo.style.display = 'block';

  meetingTitle.textContent = metadata.title || 'Google Meet';

  if (metadata.startTime) {
    const date = new Date(metadata.startTime);
    meetingTime.textContent = date.toLocaleString();
  }

  if (metadata.attendees && metadata.attendees.length > 0) {
    const attendeeText = metadata.attendees.length === 1
      ? '1 attendee'
      : `${metadata.attendees.length} attendees`;
    meetingAttendees.textContent = attendeeText;
  } else {
    meetingAttendees.textContent = '';
  }
}

/**
 * Setup event listeners
 */
function setupEventListeners() {
  startBtn.addEventListener('click', handleStartCapture);
  stopBtn.addEventListener('click', handleStopCapture);
  dismissError.addEventListener('click', hideError);

  // Listen for messages from background
  chrome.runtime.onMessage.addListener((message) => {
    console.log('[Popup] Message received:', message);

    if (message.type === 'caption_added') {
      addCaption(message.caption);
    }
    else if (message.type === 'WS_CONNECTED') {
      updateConnectionStatus(true);
    }
    else if (message.type === 'WS_DISCONNECTED') {
      updateConnectionStatus(false);
    }
  });
}

/**
 * Handle start capture
 */
async function handleStartCapture() {
  if (!meetingMetadata) {
    showError('Meeting information not available');
    return;
  }

  try {
    startBtn.disabled = true;
    showLoading('Starting capture...');

    const response = await sendMessageToBackground({
      type: 'START_SESSION',
      metadata: meetingMetadata
    });

    hideLoading();

    if (response.success) {
      currentSession = {
        sessionId: response.sessionId,
        metadata: meetingMetadata,
        captions: []
      };

      showCapturingState();
      showSuccess('Capture started! Speak to see captions.', false);

      // Clear success message after 2 seconds
      setTimeout(hideSuccess, 2000);
    } else {
      throw new Error(response.error || 'Failed to start session');
    }

  } catch (err) {
    hideLoading();
    startBtn.disabled = false;
    console.error('[Popup] Error starting capture:', err);
    showError('Failed to start capture: ' + err.message);
  }
}

/**
 * Handle stop capture
 */
async function handleStopCapture() {
  try {
    stopBtn.disabled = true;
    showLoading('Saving notes...');

    // End session
    const endResponse = await sendMessageToBackground({
      type: 'END_SESSION'
    });

    if (!endResponse.success) {
      throw new Error(endResponse.error || 'Failed to end session');
    }

    // Finalize meeting (generate summary and create doc)
    const finalizeResponse = await sendMessageToBackground({
      type: 'FINALIZE_MEETING',
      sessionId: currentSession.sessionId
    });

    hideLoading();

    if (finalizeResponse.success) {
      const result = finalizeResponse.result;

      // Show success message with doc link
      successText.textContent = `Meeting notes created with ${captions.length} captions`;
      docLink.href = result.docUrl;
      docLink.style.display = 'flex';

      showSuccess('Meeting notes saved successfully!', true);

      // Reset state
      currentSession = null;
      captions = [];

      showReadyState();
    } else {
      throw new Error(finalizeResponse.error || 'Failed to create meeting notes');
    }

  } catch (err) {
    hideLoading();
    stopBtn.disabled = false;
    console.error('[Popup] Error stopping capture:', err);
    showError('Failed to save notes: ' + err.message);
  }
}

/**
 * Add caption to UI
 */
function addCaption(caption) {
  captions.push(caption);

  // Remove "no captions" message
  const noCaptions = transcriptContent.querySelector('.no-captions');
  if (noCaptions) {
    noCaptions.remove();
  }

  // Create caption element
  const captionEl = document.createElement('div');
  captionEl.className = 'caption-item';

  const speakerEl = document.createElement('div');
  speakerEl.className = 'caption-speaker';
  speakerEl.textContent = caption.speaker;

  const textEl = document.createElement('div');
  textEl.className = 'caption-text';
  textEl.textContent = caption.text;

  const timeEl = document.createElement('div');
  timeEl.className = 'caption-time';
  const time = new Date(caption.timestamp);
  timeEl.textContent = time.toLocaleTimeString();

  captionEl.appendChild(speakerEl);
  captionEl.appendChild(textEl);
  captionEl.appendChild(timeEl);

  transcriptContent.appendChild(captionEl);

  // Auto-scroll to bottom
  transcriptContent.scrollTop = transcriptContent.scrollHeight;

  // Update count
  captionCount.textContent = `${captions.length} caption${captions.length !== 1 ? 's' : ''}`;
}

/**
 * Show capturing state
 */
function showCapturingState() {
  startBtn.style.display = 'none';
  stopBtn.style.display = 'flex';
  stopBtn.disabled = false;
  transcriptContainer.style.display = 'flex';
  statusIndicator.classList.add('capturing');
  statusIndicator.querySelector('.status-text').textContent = 'Capturing';

  // Display existing captions
  transcriptContent.innerHTML = '';
  if (captions.length === 0) {
    transcriptContent.innerHTML = '<div class="no-captions">Waiting for captions...</div>';
  } else {
    captions.forEach(caption => addCaption(caption));
  }
}

/**
 * Show ready state
 */
function showReadyState() {
  startBtn.style.display = 'flex';
  startBtn.disabled = false;
  stopBtn.style.display = 'none';
  transcriptContainer.style.display = 'none';
  statusIndicator.classList.remove('capturing');
  statusIndicator.querySelector('.status-text').textContent = 'Ready';
}

/**
 * Update connection status
 */
function updateConnectionStatus(connected) {
  if (connected) {
    connectionStatus.classList.add('connected');
    connectionStatus.classList.remove('disconnected');
    wsStatus.textContent = 'Connected';
  } else {
    connectionStatus.classList.remove('connected');
    connectionStatus.classList.add('disconnected');
    wsStatus.textContent = 'Disconnected';
  }
}

/**
 * Show loading
 */
function showLoading(text) {
  loadingText.textContent = text;
  loading.style.display = 'block';
}

/**
 * Hide loading
 */
function hideLoading() {
  loading.style.display = 'none';
}

/**
 * Show success message
 */
function showSuccess(text, persistent = false) {
  if (!persistent) {
    const tempSuccess = document.createElement('div');
    tempSuccess.className = 'success-message';
    tempSuccess.style.padding = '12px';
    tempSuccess.innerHTML = `<p style="margin: 0; color: #34A853; font-size: 13px;">${text}</p>`;
    document.querySelector('.controls').appendChild(tempSuccess);

    setTimeout(() => {
      tempSuccess.remove();
    }, 2000);
  } else {
    successMessage.style.display = 'block';
    meetingInfo.style.display = 'none';
  }
}

/**
 * Hide success message
 */
function hideSuccess() {
  successMessage.style.display = 'none';
}

/**
 * Show error
 */
function showError(text) {
  errorText.textContent = text;
  errorMessage.style.display = 'block';
}

/**
 * Hide error
 */
function hideError() {
  errorMessage.style.display = 'none';
}

/**
 * Send message to background script
 */
function sendMessageToBackground(message) {
  return new Promise((resolve, reject) => {
    chrome.runtime.sendMessage(message, (response) => {
      if (chrome.runtime.lastError) {
        reject(new Error(chrome.runtime.lastError.message));
      } else {
        resolve(response);
      }
    });
  });
}

// Initialize on load
initialize();
