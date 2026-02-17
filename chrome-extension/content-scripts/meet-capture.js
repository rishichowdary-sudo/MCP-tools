/**
 * Google Meet Caption Capture — Clean Rewrite
 *
 * Core problem solved: Previous version read the ENTIRE caption container's
 * textContent which accumulated ALL speakers' text causing duplication.
 *
 * New approach:
 * - Track each INDIVIDUAL speaker block as a separate DOM element
 * - Read ONLY that block's text, never the parent container
 * - Debounce per-element: wait 2s after last change before saving
 * - Finalize immediately when the block disappears from DOM
 * - Proper speaker name extraction from dedicated name element within each block
 */

console.log('[Meet Capture] Content script v2 loaded');

// ── State ──────────────────────────────────────────────────────────────────
let isCapturing = false;
let mainObserver = null;
let pollInterval = null;

// WeakMap so entries are GC'd when elements are removed from DOM
const blockState = new WeakMap();
// elementRef → { speaker, lastText, lastSentText, timer }

// Track sent captions to avoid exact duplicates across elements
const recentlySent = [];
const MAX_RECENT = 50;
const DEDUPE_WINDOW_MS = 20000;

const DEBOUNCE_MS = 1800; // ms of silence before a caption is "done"
const MIN_LENGTH  = 3;    // ignore ultra-short fragments only

// ── Caption block detection ────────────────────────────────────────────────

/**
 * Google Meet renders captions in a scrollable list.
 * Each LIST ITEM inside that list = one speaker's utterance.
 * We scan for those items every time mutations fire.
 *
 * Selectors are tried in order; first match wins.
 * These are the known patterns as of early 2026:
 */
const CONTAINER_SELECTORS = [
  '[jsname="tgaKEf"]',   // Outer caption panel
  '[jsname="dsyhDe"]',   // Inner list
  '.iOzk7',
  '[data-is-caption="true"]',
  '.a4cQT',
];

// Within a caption block, these selectors find speaker name
const SPEAKER_SELECTORS = [
  '[jsname="r4nke"]',    // Most reliable as of 2025-2026
  '.NWpY1d',
  '.zs7s8d',
  '[class*="speaker" i]',
  '[class*="name" i]',
];

// Within a caption block, these selectors find the spoken text
const TEXT_SELECTORS = [
  '[jsname="YSxPC"]',
  '[jsname="K4s0"]',
  '.a4cQT',
  '[class*="caption-text" i]',
  '[class*="transcript" i]',
];

function getCaptionContainer() {
  for (const sel of CONTAINER_SELECTORS) {
    const el = document.querySelector(sel);
    if (el) return el;
  }
  return null;
}

/**
 * Returns all current caption blocks (individual speaker utterance elements)
 * inside the caption container.
 */
function getCaptionBlocks(container) {
  if (!container) return [];

  // Strategy 1: Direct children that contain both speaker + text sub-elements
  const children = Array.from(container.children);
  if (children.length > 0) {
    // Filter to children that look like caption blocks
    const blocks = children.filter(el => {
      const hasSpeaker = SPEAKER_SELECTORS.some(s => el.querySelector(s));
      const hasText    = TEXT_SELECTORS.some(s => el.querySelector(s));
      return hasSpeaker || hasText;
    });
    if (blocks.length > 0) return blocks;
    // Fall back to ALL children
    return children;
  }

  // Strategy 2: Query for known block-level elements
  return Array.from(container.querySelectorAll(':scope > div, :scope > li'));
}

// ── Text extraction ────────────────────────────────────────────────────────

function extractFromBlock(block) {
  // --- Speaker name ---
  let speaker = '';
  for (const sel of SPEAKER_SELECTORS) {
    const el = block.querySelector(sel);
    if (el) {
      speaker = el.textContent.trim().replace(/:$/, '').trim();
      if (speaker) break;
    }
  }

  // --- Spoken text ---
  let text = '';
  for (const sel of TEXT_SELECTORS) {
    const el = block.querySelector(sel);
    if (el) {
      text = el.textContent.trim();
      if (text) break;
    }
  }

  // Fallback: if neither selector worked, try to split on first ":" in full text
  if (!text && !speaker) {
    const full = block.textContent.trim();
    const colonIdx = full.indexOf(':');
    if (colonIdx > 0 && colonIdx < 40) {
      speaker = full.slice(0, colonIdx).trim();
      text    = full.slice(colonIdx + 1).trim();
    } else {
      text = full;
    }
  }

  // If we have text but no speaker, try to infer from the block's aria/title attrs
  if (!speaker) {
    speaker = block.getAttribute('aria-label') || block.getAttribute('title') || 'You';
    // Strip common suffixes
    speaker = speaker.replace(/['']s caption/i, '').trim();
  }

  // Clean speaker
  speaker = speaker.replace(/\s+/g, ' ').trim() || 'You';

  // Clean text
  text = cleanText(text);

  return { speaker, text };
}

function cleanText(raw) {
  return raw
    .replace(/\s+/g, ' ')
    .replace(/^[.,\s]+/, '')
    .trim();
}

// ── Duplicate detection ────────────────────────────────────────────────────

function alreadySent(speaker, text) {
  const now = Date.now();
  for (let i = recentlySent.length - 1; i >= 0; i--) {
    if (now - recentlySent[i].ts > DEDUPE_WINDOW_MS) {
      recentlySent.splice(i, 1);
    }
  }
  return recentlySent.some(r => r.speaker === speaker && r.text === text);
}

function markSent(speaker, text) {
  recentlySent.push({ speaker, text, ts: Date.now() });
  if (recentlySent.length > MAX_RECENT) recentlySent.shift();
}

// ── Sending ────────────────────────────────────────────────────────────────

function sendCaption(speaker, text) {
  if (!text || text.length < MIN_LENGTH) return;
  if (alreadySent(speaker, text)) return;

  // Reject pure noise
  if (/^[^a-zA-Z]*$/.test(text)) return;

  markSent(speaker, text);

  const caption = { speaker, text, timestamp: new Date().toISOString() };
  console.log('[Meet Capture] ✅', speaker + ':', text.slice(0, 60));

  chrome.runtime.sendMessage({ type: 'CAPTION_CAPTURED', caption }).catch(() => {});
}

// ── Per-block tracking ─────────────────────────────────────────────────────

function processBlock(block) {
  if (!isCapturing) return;

  const { speaker, text } = extractFromBlock(block);
  if (!text || text.length < MIN_LENGTH) return;

  let state = blockState.get(block);
  if (!state) {
    state = { speaker, lastText: '', lastSentText: '', timer: null };
    blockState.set(block, state);
  }

  // Update speaker (it might change if element is reused)
  if (speaker) state.speaker = speaker;

  // No change since last check — nothing to do
  if (text === state.lastText) return;
  state.lastText = text;

  // Reset debounce timer
  if (state.timer) clearTimeout(state.timer);
  state.timer = setTimeout(() => {
    // Check if the block is still in the DOM
    const inDom = document.contains(block);
    const finalText = inDom ? extractFromBlock(block).text : state.lastText;

    if (finalText && finalText !== state.lastSentText && finalText.length >= MIN_LENGTH) {
      sendCaption(state.speaker, finalText);
      state.lastSentText = finalText;
    }
    state.timer = null;
  }, DEBOUNCE_MS);
}

/**
 * Called when blocks are REMOVED from the DOM — finalize them immediately.
 */
function finalizeBlock(block) {
  const state = blockState.get(block);
  if (!state) return;

  if (state.timer) {
    clearTimeout(state.timer);
    state.timer = null;
  }

  const text = state.lastText;
  if (text && text !== state.lastSentText && text.length >= MIN_LENGTH) {
    sendCaption(state.speaker, text);
    state.lastSentText = text;
  }
}

// ── MutationObserver ───────────────────────────────────────────────────────

function handleMutations(mutations) {
  if (!isCapturing) return;

  const container = getCaptionContainer();
  if (!container) return;

  let needsBlockScan = false;

  for (const m of mutations) {
    // New caption blocks added
    if (m.addedNodes.length > 0) needsBlockScan = true;

    // Caption blocks removed — finalize them
    for (const node of m.removedNodes) {
      if (node.nodeType === Node.ELEMENT_NODE) {
        finalizeBlock(node);
        // Also check children of removed node
        node.querySelectorAll && node.querySelectorAll('*').forEach(child => finalizeBlock(child));
      }
    }

    // Text updated inside an existing block — find the block and process it
    if (m.type === 'characterData' || m.type === 'childList') {
      let el = m.target.nodeType === Node.ELEMENT_NODE ? m.target : m.target.parentElement;
      // Walk up until we find a direct child of the container (= a caption block)
      while (el && el.parentElement !== container) {
        el = el.parentElement;
      }
      if (el && el.parentElement === container) {
        processBlock(el);
        needsBlockScan = false; // Already processed the right block
      } else {
        needsBlockScan = true;
      }
    }
  }

  if (needsBlockScan) {
    getCaptionBlocks(container).forEach(processBlock);
  }
}

// ── Polling fallback ───────────────────────────────────────────────────────
// Some Meet versions update DOM in ways MutationObserver misses.
// A light poll every 800ms catches those.

function pollBlocks() {
  if (!isCapturing) return;
  const container = getCaptionContainer();
  if (container) getCaptionBlocks(container).forEach(processBlock);
}

// ── Start / Stop ───────────────────────────────────────────────────────────

function startCapture() {
  if (isCapturing) return;
  isCapturing = true;

  const container = getCaptionContainer();
  if (!container) {
    console.warn('[Meet Capture] Caption container not found — retrying in 2s. Make sure captions (CC) are ON.');
    setTimeout(startCapture, 2000);
    isCapturing = false;
    return;
  }

  mainObserver = new MutationObserver(handleMutations);
  mainObserver.observe(container, { childList: true, subtree: true, characterData: true });

  pollInterval = setInterval(pollBlocks, 800);

  console.log('[Meet Capture] Started — container found:', container.tagName, container.className.slice(0, 40));
  chrome.runtime.sendMessage({ type: 'CAPTURE_STARTED' }).catch(() => {});
}

function stopCapture() {
  isCapturing = false;

  if (mainObserver) { mainObserver.disconnect(); mainObserver = null; }
  if (pollInterval) { clearInterval(pollInterval); pollInterval = null; }

  recentlySent.length = 0;
  console.log('[Meet Capture] Stopped');
  chrome.runtime.sendMessage({ type: 'CAPTURE_STOPPED' }).catch(() => {});
}

// ── Message listener ───────────────────────────────────────────────────────

chrome.runtime.onMessage.addListener((msg, _sender, respond) => {
  if (msg.type === 'START_CAPTURE') { startCapture(); respond({ success: true }); }
  else if (msg.type === 'STOP_CAPTURE') { stopCapture(); respond({ success: true }); }
  else if (msg.type === 'GET_STATUS') {
    respond({ isCapturing, hasCaptionContainer: !!getCaptionContainer() });
  }
  return true;
});

// ── Meeting detection ──────────────────────────────────────────────────────

function getMeetingCode() {
  const m = window.location.href.match(/meet\.google\.com\/([a-z0-9-]+)/);
  return m ? m[1] : null;
}

function getMeetingTitle() {
  const selectors = ['[data-meeting-title]', '.u6vdEc', '[role="heading"]', 'h1'];
  for (const sel of selectors) {
    const el = document.querySelector(sel);
    if (!el) continue;
    const t = (el.getAttribute('data-meeting-title') || el.textContent || '').trim();
    if (t && t.length > 2 && !t.toLowerCase().includes('google meet')) return t;
  }
  const parts = document.title.split(' - ');
  if (parts.length > 1) return parts[0].trim();
  return null;
}

// Send meeting info on load
const meetingCode = getMeetingCode();
if (meetingCode) {
  chrome.runtime.sendMessage({
    type: 'MEETING_DETECTED',
    meetingCode,
    meetingTitle: getMeetingTitle(),
    url: window.location.href
  }).catch(() => {});
}

// Auto-start if setting is on
chrome.storage.local.get(['autoCapture'], result => {
  if (result.autoCapture) setTimeout(startCapture, 3000);
});
