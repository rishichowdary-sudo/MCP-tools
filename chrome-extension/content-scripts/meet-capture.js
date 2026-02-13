/**
 * Google Meet Caption Capture Content Script
 * Extracts live captions from Google Meet and sends to background script
 */

console.log('[Meet Capture] Content script loaded');

// Configuration
const CAPTION_SELECTORS = [
  '[jsname="dsyhDe"]',              // Primary caption container
  '.iOzk7',                         // Alternative container
  '[data-is-caption="true"]',       // Caption marker attribute
  '.a4cQT',                         // Caption text container
  '.TBMuR'                          // Speaker name container
];

let observer = null;
let captionContainer = null;
let isCapturing = false;
let processedCaptions = new Set(); // Prevent duplicate captures
let activeCaptionElement = null; // Track the currently updating caption element
let captionStabilizeTimer = null; // Timer to wait for caption to stabilize
let lastCaptionText = ''; // Track last captured text
let currentUserName = null; // Store current user's name
const CAPTION_STABILIZE_DELAY = 1500; // Wait 1.5s after last change before capturing

/**
 * Find the caption container in the DOM
 */
function findCaptionContainer() {
  for (const selector of CAPTION_SELECTORS) {
    const element = document.querySelector(selector);
    if (element) {
      console.log('[Meet Capture] Found caption container:', selector);
      return element;
    }
  }

  // Fallback: look for elements with caption-related classes
  const captionElements = document.querySelectorAll('[class*="caption"], [class*="subtitle"]');
  if (captionElements.length > 0) {
    console.log('[Meet Capture] Found caption element via fallback');
    return captionElements[0].parentElement;
  }

  return null;
}

/**
 * Get current user's name from Google Meet page
 */
function getCurrentUserName() {
  if (currentUserName) return currentUserName;

  // Strategy 1: Get from participant list (most reliable)
  try {
    // Look for the participant panel
    const participantItems = document.querySelectorAll('[data-self-name], [data-requested-participant-id]');
    for (const item of participantItems) {
      const selfName = item.getAttribute('data-self-name');
      if (selfName && selfName.length > 0 && selfName !== 'You') {
        currentUserName = selfName;
        console.log('[Meet Capture] Detected user name from participant:', currentUserName);
        return currentUserName;
      }
    }
  } catch (e) {}

  // Strategy 2: Get from Google account button
  try {
    const accountButtons = document.querySelectorAll('[aria-label*="Google Account"]');
    for (const btn of accountButtons) {
      const label = btn.getAttribute('aria-label');
      if (label && label.includes(':')) {
        const parts = label.split(':');
        if (parts.length > 1) {
          const namePart = parts[1].trim().split(/[(\s]/)[0]; // Get first name before space or (
          if (namePart && namePart.length > 1 && !namePart.includes('alarm')) {
            currentUserName = namePart;
            console.log('[Meet Capture] Detected user name from account:', currentUserName);
            return currentUserName;
          }
        }
      }
    }
  } catch (e) {}

  // Strategy 3: Get from profile image alt text
  try {
    const profileImgs = document.querySelectorAll('[data-disable-tooltip] img, [aria-label*="profile" i] img');
    for (const img of profileImgs) {
      const alt = img.getAttribute('alt') || img.getAttribute('aria-label');
      if (alt && alt.length > 0 && alt !== 'You' && !alt.includes('alarm')) {
        // Extract just the name part
        const name = alt.split(/[(\s]/)[0];
        if (name && name.length > 1) {
          currentUserName = name;
          console.log('[Meet Capture] Detected user name from profile:', currentUserName);
          return currentUserName;
        }
      }
    }
  } catch (e) {}

  console.log('[Meet Capture] Could not detect user name, using default');
  return 'You'; // Default fallback
}

/**
 * Extract speaker and text from caption element
 */
function extractCaptionFromElement(element) {
  let speaker = getCurrentUserName(); // Default to current user
  let text = '';

  // Method 1: Look for specific structure (speaker + text containers)
  const speakerEl = element.querySelector('.TBMuR, [class*="speaker"], [class*="name"]');
  const textEl = element.querySelector('.a4cQT, [class*="caption-text"], [class*="text"]');

  if (speakerEl && textEl) {
    const speakerText = speakerEl.textContent.trim().replace(':', '');
    speaker = speakerText || speaker;
    text = textEl.textContent.trim();
  }
  // Method 2: Parse combined text (format: "Speaker: Caption text")
  else if (element.textContent.includes(':')) {
    const parts = element.textContent.split(':');
    if (parts.length >= 2) {
      const speakerPart = parts[0].trim();
      // Only use if it looks like a name (not "You" alone)
      if (speakerPart && speakerPart.length > 1 && speakerPart !== 'You') {
        speaker = speakerPart;
      }
      text = parts.slice(1).join(':').trim();
    }
  }
  // Method 3: Just caption text (no speaker identified)
  else {
    text = element.textContent.trim();
  }

  // Clean up speaker name
  if (speaker === 'You') {
    speaker = getCurrentUserName();
  }

  return { speaker, text };
}

/**
 * Handle caption element updates (the robust way)
 */
function handleCaptionElement(element) {
  if (!element || !isCapturing) return;

  // Extract current caption data
  const { speaker, text } = extractCaptionFromElement(element);

  // Ignore empty or very short captions (likely fragments)
  if (!text || text.length < 3) {
    return;
  }

  // Check if this is a new caption or an update to existing one
  const isSameAsLast = text === lastCaptionText;
  if (isSameAsLast) {
    return; // No change, ignore
  }

  // Check if this text is a substring/prefix of last text (likely old fragment)
  if (lastCaptionText && lastCaptionText.includes(text) && lastCaptionText.length > text.length) {
    console.log('[Meet Capture] Ignoring fragment (subset of current):', text);
    return; // This is an old fragment of the current caption, ignore
  }

  // Check if this is the same element still being updated
  const isSameElement = element === activeCaptionElement;

  if (isSameElement) {
    // Element is still being updated - reset stabilize timer
    console.log('[Meet Capture] Caption updating:', text.substring(0, 50) + '...');
  } else {
    // New caption element detected
    // ONLY finalize previous if it was significantly different (not a fragment)
    if (activeCaptionElement && lastCaptionText && lastCaptionText.length >= 10) {
      // Only finalize if the previous caption was substantial
      console.log('[Meet Capture] New element, finalizing previous caption');
      finalizePreviousCaption();
    } else if (activeCaptionElement && lastCaptionText) {
      // Previous caption was too short, probably a fragment - don't send it
      console.log('[Meet Capture] Discarding short previous caption:', lastCaptionText);
      if (captionStabilizeTimer) {
        clearTimeout(captionStabilizeTimer);
        captionStabilizeTimer = null;
      }
    }

    // Track new element
    activeCaptionElement = element;
    console.log('[Meet Capture] New caption element detected');
  }

  // Update last text
  lastCaptionText = text;

  // Clear existing timer
  if (captionStabilizeTimer) {
    clearTimeout(captionStabilizeTimer);
  }

  // Wait for caption to stabilize (no changes for CAPTION_STABILIZE_DELAY)
  captionStabilizeTimer = setTimeout(() => {
    finalizeCaption(element, speaker, text);
  }, CAPTION_STABILIZE_DELAY);
}

/**
 * Finalize previous caption before moving to next
 */
function finalizePreviousCaption() {
  if (activeCaptionElement && lastCaptionText) {
    const { speaker } = extractCaptionFromElement(activeCaptionElement);
    finalizeCaption(activeCaptionElement, speaker, lastCaptionText);
  }
}

/**
 * Finalize and send caption (after it's stabilized)
 */
function finalizeCaption(element, speaker, text) {
  // Ignore if too short (likely a fragment that slipped through)
  if (text.length < 5) {
    console.log('[Meet Capture] Discarding short caption:', text);
    return;
  }

  // Check for duplicates
  const captionId = `${speaker}:${text}`;
  if (processedCaptions.has(captionId)) {
    console.log('[Meet Capture] Duplicate caption, skipping');
    return;
  }

  // Check if this is a fragment of something already sent
  for (const processed of processedCaptions) {
    if (processed.includes(text) && processed.length > text.length) {
      console.log('[Meet Capture] This is a fragment of already sent caption, skipping');
      return;
    }
  }

  processedCaptions.add(captionId);

  // Clean up old entries (keep last 20)
  if (processedCaptions.size > 20) {
    const iterator = processedCaptions.values();
    processedCaptions.delete(iterator.next().value);
  }

  const caption = {
    speaker,
    text,
    timestamp: new Date().toISOString()
  };

  console.log('[Meet Capture] ✅ SENDING caption:', caption);

  // Send to background script
  chrome.runtime.sendMessage({
    type: 'CAPTION_CAPTURED',
    caption: caption
  }).catch(err => {
    console.error('[Meet Capture] Error sending caption:', err);
  });

  // Clear tracking for next caption
  if (element === activeCaptionElement) {
    activeCaptionElement = null;
    lastCaptionText = '';
  }
}

/**
 * Handle caption mutations (optimized for real-time updates)
 */
function handleMutations(mutations) {
  if (!isCapturing) return;

  for (const mutation of mutations) {
    // Handle added nodes (new caption elements)
    if (mutation.type === 'childList' && mutation.addedNodes.length > 0) {
      for (const node of mutation.addedNodes) {
        if (node.nodeType === Node.ELEMENT_NODE) {
          handleCaptionElement(node);
        }
      }
    }

    // Handle text changes (caption being updated)
    if (mutation.type === 'characterData' || mutation.type === 'childList') {
      const target = mutation.target;
      const element = target.nodeType === Node.ELEMENT_NODE ? target : target.parentElement;
      if (element) {
        handleCaptionElement(element);
      }
    }
  }
}

/**
 * Start observing captions
 */
function startObserving() {
  if (observer) {
    console.log('[Meet Capture] Observer already running');
    return;
  }

  captionContainer = findCaptionContainer();

  if (!captionContainer) {
    console.warn('[Meet Capture] Caption container not found. Make sure captions are enabled in Meet.');

    // Retry after 2 seconds (captions might not be loaded yet)
    setTimeout(startObserving, 2000);
    return;
  }

  observer = new MutationObserver(handleMutations);
  observer.observe(captionContainer, {
    childList: true,
    subtree: true,
    characterData: true,  // Watch for text changes
    characterDataOldValue: true
  });

  isCapturing = true;
  console.log('[Meet Capture] Started observing captions');

  // Notify background script
  chrome.runtime.sendMessage({
    type: 'CAPTURE_STARTED'
  });
}

/**
 * Stop observing captions
 */
function stopObserving() {
  if (observer) {
    observer.disconnect();
    observer = null;
  }

  isCapturing = false;
  processedCaptions.clear();
  console.log('[Meet Capture] Stopped observing captions');

  // Notify background script
  chrome.runtime.sendMessage({
    type: 'CAPTURE_STOPPED'
  });
}

/**
 * Listen for messages from background script
 */
chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  console.log('[Meet Capture] Received message:', message);

  if (message.type === 'START_CAPTURE') {
    startObserving();
    sendResponse({ success: true });
  }
  else if (message.type === 'STOP_CAPTURE') {
    stopObserving();
    sendResponse({ success: true });
  }
  else if (message.type === 'GET_STATUS') {
    sendResponse({
      isCapturing,
      hasCaptionContainer: captionContainer !== null
    });
  }

  return true; // Keep channel open for async response
});

/**
 * Extract meeting code from URL
 */
function getMeetingCode() {
  const url = window.location.href;
  const match = url.match(/meet\.google\.com\/([a-z-]+)/);
  return match ? match[1] : null;
}

/**
 * Send meeting info to background on load
 */
function sendMeetingInfo() {
  const meetingCode = getMeetingCode();

  if (meetingCode) {
    chrome.runtime.sendMessage({
      type: 'MEETING_DETECTED',
      meetingCode: meetingCode,
      url: window.location.href
    });
  }
}

// Initialize
sendMeetingInfo();

// Auto-start if user has enabled auto-capture
chrome.storage.local.get(['autoCapture'], (result) => {
  if (result.autoCapture) {
    console.log('[Meet Capture] Auto-capture enabled, starting observation');
    setTimeout(startObserving, 3000); // Wait for page to fully load
  }
});
