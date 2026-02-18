/**
 * Google Meet Caption Capture — v3 (Robust Discovery)
 *
 * v2 relied on hard-coded jsname / class selectors that Google changed
 * when they shipped scrollable captions (Feb 2025) and customizable
 * caption styling (Jun 2025).
 *
 * v3 keeps the old selectors as quick-hit hints but adds heuristic-based
 * discovery so it works even when Google changes their DOM again.
 *
 * Approach:
 *  1. Try known selectors first (fast path).
 *  2. If they all miss, scan the DOM for elements that look like captions:
 *     - Fixed/absolute-positioned overlay near the bottom of the viewport
 *     - Contains rapidly-changing text with speaker-like patterns
 *  3. Once found, attach MutationObserver + polling exactly like v2.
 *  4. A body-level MutationObserver continuously watches for the caption
 *     panel appearing (user toggles CC on/off during a call).
 */

console.log('[Meet Capture] Content script v3 loaded');

// ── State ──────────────────────────────────────────────────────────────────
let isCapturing = false;
let mainObserver = null;
let discoveryObserver = null;
let pollInterval = null;
let discoveryInterval = null;
let lastKnownContainer = null;

// WeakMap so entries are GC'd when elements are removed from DOM
const blockState = new WeakMap();

// Track sent captions to avoid exact duplicates across elements
const recentlySent = [];
const MAX_RECENT = 50;
const DEDUPE_WINDOW_MS = 20000;

const DEBOUNCE_MS = 1800;
const MIN_LENGTH = 3;
const UI_NOISE_PATTERNS = [
  /arrow_downward/ig,
  /jump to bottom/ig,
  /^captions$/ig,
  /turn on captions/ig,
  /captions are turned off/ig,
  /captions are on/ig,
  /you turned on captions/ig,
];

// ── Selector banks ─────────────────────────────────────────────────────────
// Kept from v2 as "known-good hints". They will be tried first.

const CONTAINER_SELECTORS = [
  // 2024-era selectors (may still work for some users)
  '[jsname="tgaKEf"]',
  '[jsname="dsyhDe"]',
  '.iOzk7',
  '[data-is-caption="true"]',
  '.a4cQT',
  // 2025-era scrollable caption selectors — these are educated guesses
  '[jsname="B0czdc"]',
  '[jsname="YPqjbf"]',
  // Generic: any element with role=region that sits at the bottom
  '[role="region"][aria-live]',
  '[aria-live="polite"]',
  '[aria-live="assertive"]',
];

const SPEAKER_SELECTORS = [
  '[jsname="r4nke"]',
  '.NWpY1d',
  '.zs7s8d',
  '[class*="speaker" i]',
  '[class*="name" i]',
  // 2025+ additions
  '[data-sender-name]',
  '[data-speaker]',
];

const TEXT_SELECTORS = [
  // 2026-era generic fallbacks
  '[jsname="YSxPC"]',
  '[jsname="K4s0"]',
  '.a4cQT',
  '.iTTPOb',
  '.VbkSUe', // New common text container
  '.cn87Jb', // Another text container variant
  '[class*="caption-text" i]',
  '[class*="transcript" i]',
  // Generic fallback
  'span',
  'div[jsname]'
];

// ── Heuristic caption container discovery ──────────────────────────────────

/**
 * Score an element on how likely it is to be the caption overlay.
 * Higher = more likely.
 */
function scoreCaptionCandidate(el) {
  let score = 0;
  const rect = el.getBoundingClientRect();
  const style = window.getComputedStyle(el);

  // Must be visible
  if (rect.width === 0 || rect.height === 0) return -1;
  if (style.display === 'none' || style.visibility === 'hidden') return -1;
  // Skip tiny elements
  if (rect.width < 100 || rect.height < 20) return -1;

  // Positioned near the bottom of the viewport
  const viewH = window.innerHeight;
  if (rect.bottom > viewH * 0.6) score += 3;
  if (rect.bottom > viewH * 0.75) score += 2;

  // Fixed or absolute positioning (overlaid on video)
  if (style.position === 'fixed' || style.position === 'absolute') score += 3;

  // Contains text that looks conversational (has letters, spaces)
  const text = el.textContent || '';
  if (text.length > 5 && /[a-zA-Z]{2,}/.test(text)) score += 2;

  // Has child elements (speaker blocks)
  if (el.children.length >= 1) score += 1;
  if (el.children.length >= 2) score += 1;

  // aria-live is a strong caption signal
  if (el.getAttribute('aria-live')) score += 5;
  if (el.getAttribute('role') === 'region') score += 2;

  // Contains a colon (speaker: text pattern)
  if (text.includes(':')) score += 2;

  // Specific check for "You" or known speaker patterns
  if (/^[A-Z][a-z]+(\s[A-Z][a-z]+)*:/.test(text)) score += 3;

  // Penalty: if it looks like a toolbar, chat, or participant list
  const cl = (el.className || '').toLowerCase();
  const id = (el.id || '').toLowerCase();
  if (cl.includes('toolbar') || cl.includes('chat') || cl.includes('participant')) score -= 5;
  if (id.includes('toolbar') || id.includes('chat') || id.includes('participant')) score -= 5;

  // Penalty: too large (probably the whole page)
  if (rect.width > viewH && rect.height > viewH) score -= 3;

  return score;
}

/**
 * Walk the DOM looking for the most likely caption container.
 */
function discoverCaptionContainer() {
  // Strategy 1: Elements with aria-live (accessibility-driven caption regions)
  const ariaLive = document.querySelectorAll('[aria-live="polite"], [aria-live="assertive"]');
  let best = null;
  let bestScore = 0;

  for (const el of ariaLive) {
    const s = scoreCaptionCandidate(el);
    if (s > bestScore) { bestScore = s; best = el; }
  }

  // Strategy 2: Fixed/absolute-positioned elements near the bottom
  if (!best || bestScore < 5) {
    const allDivs = document.querySelectorAll('div, section, aside');
    for (const el of allDivs) {
      // Quick pre-filter: skip tiny and invisible
      if (!el.offsetParent && el !== document.body) continue;
      const s = scoreCaptionCandidate(el);
      if (s > bestScore) { bestScore = s; best = el; }
    }
  }

  if (best && bestScore >= 5) {
    console.log(`[Meet Capture] 🔍 Discovered caption container via heuristic (score=${bestScore}):`,
      best.tagName, best.className?.slice?.(0, 50) || '', best.id || '');
    return best;
  }

  return null;
}

function getCaptionContainer() {
  // Fast path: check if our last-known container is still in the DOM and visible
  if (lastKnownContainer && document.contains(lastKnownContainer)) {
    const rect = lastKnownContainer.getBoundingClientRect();
    if (rect.width > 0 && rect.height > 0) return lastKnownContainer;
  }

  // Try known selectors
  for (const sel of CONTAINER_SELECTORS) {
    try {
      const els = document.querySelectorAll(sel);
      for (const el of els) {
        if (looksLikeCaptionContainer(el)) {
          lastKnownContainer = el;
          console.log(`[Meet Capture] Found container via selector: ${sel}`);
          return el;
        }
      }
    } catch { /* invalid selector, skip */ }
  }

  // Heuristic discovery fallback
  const discovered = discoverCaptionContainer();
  if (discovered) {
    lastKnownContainer = discovered;
    return discovered;
  }

  return null;
}

function looksLikeCaptionContainer(el) {
  if (!el || !el.querySelectorAll) return false;
  const rect = el.getBoundingClientRect();
  if (rect.width === 0 || rect.height === 0) return false;

  // Check if it has speaker + text sub-elements (known structure)
  const hasSpeaker = SPEAKER_SELECTORS.some(s => { try { return !!el.querySelector(s); } catch { return false; } });
  const hasText = TEXT_SELECTORS.some(s => { try { return !!el.querySelector(s); } catch { return false; } });
  if (hasSpeaker && hasText) return true;

  // Check if text content looks like captions
  const text = (el.textContent || '').trim();
  if (text.length > 10 && text.length < 2000 && /[a-zA-Z]{3,}/.test(text)) {
    // Has child elements that could be speaker blocks
    if (el.children.length >= 1) return true;
  }

  return false;
}

// ── Caption block detection ────────────────────────────────────────────────

function getCaptionBlocks(container) {
  if (!container) return [];

  const children = Array.from(container.children);
  if (children.length === 0) return [];

  // Strategy 1: children that have speaker+text sub-elements
  const withStructure = children.filter(el => {
    const hasSpeaker = SPEAKER_SELECTORS.some(s => { try { return !!el.querySelector(s); } catch { return false; } });
    const hasText = TEXT_SELECTORS.some(s => { try { return !!el.querySelector(s); } catch { return false; } });
    if (hasSpeaker && hasText) return true;
    if (hasText) {
      const extracted = extractFromBlock(el);
      return !!(extracted?.text && !isUiNoiseCaption(extracted.speaker, extracted.text));
    }
    return false;
  });
  if (withStructure.length > 0) return withStructure;

  // Strategy 2: all children with meaningful text
  const withText = children.filter(el => {
    const extracted = extractFromBlock(el);
    return !!(extracted?.text && extracted.text.length >= MIN_LENGTH && !isUiNoiseCaption(extracted.speaker, extracted.text));
  });
  if (withText.length > 0) return withText;

  // Strategy 3: dig one level deeper (container > div > block)
  const grandchildren = [];
  for (const child of children) {
    for (const gc of child.children) {
      const extracted = extractFromBlock(gc);
      if (extracted?.text && extracted.text.length >= MIN_LENGTH && !isUiNoiseCaption(extracted.speaker, extracted.text)) {
        grandchildren.push(gc);
      }
    }
  }
  if (grandchildren.length > 0) return grandchildren;

  // Strategy 4: treat the container itself as a single block
  const containerText = extractFromBlock(container);
  if (containerText?.text && containerText.text.length >= MIN_LENGTH && !isUiNoiseCaption(containerText.speaker, containerText.text)) {
    return [container];
  }

  return [];
}

// ── Text extraction ────────────────────────────────────────────────────────

function extractFromBlock(block) {
  // --- Speaker name ---
  let speaker = '';
  for (const sel of SPEAKER_SELECTORS) {
    try {
      const el = block.querySelector(sel);
      if (el) {
        speaker = el.textContent.trim().replace(/:$/, '').trim();
        if (speaker) break;
      }
    } catch { /* skip invalid selectors */ }
  }

  // --- Spoken text ---
  let text = '';
  for (const sel of TEXT_SELECTORS) {
    try {
      const el = block.querySelector(sel);
      if (el) {
        // Don't use the speaker element's text as caption text
        if (speaker && el.textContent.trim() === speaker) continue;
        text = el.textContent.trim();
        if (text) break;
      }
    } catch { /* skip */ }
  }

  // --- Fallback: parse block's full textContent ---
  if (!text && !speaker) {
    const full = block.textContent.trim();
    const colonIdx = full.indexOf(':');
    if (colonIdx > 0 && colonIdx < 40) {
      speaker = full.slice(0, colonIdx).trim();
      text = full.slice(colonIdx + 1).trim();
    } else if (full.length > 0) {
      text = full;
    }
  } else if (!text) {
    // We have a speaker but no text from selectors — get text from the block
    // minus the speaker name
    const full = block.textContent.trim();
    if (speaker && full.startsWith(speaker)) {
      text = full.slice(speaker.length).replace(/^[:\s]+/, '').trim();
    } else {
      text = full;
    }
  }

  // If we have text but no speaker, try attributes
  if (!speaker) {
    speaker = block.getAttribute('data-sender-name')
      || block.getAttribute('data-speaker')
      || block.getAttribute('aria-label')
      || block.getAttribute('title')
      || '';
    speaker = speaker.replace(/['']s caption/i, '').trim();
  }

  // Clean
  speaker = speaker.replace(/\s+/g, ' ').trim() || 'You';
  text = cleanText(text);

  // Make sure we didn't accidentally include the speaker name as the text
  if (text === speaker) text = '';

  return { speaker, text };
}

function cleanText(raw) {
  return raw
    .replace(/arrow_downward/ig, ' ')
    .replace(/\s+/g, ' ')
    .replace(/^[.,\s]+/, '')
    .trim();
}

function isUiNoiseCaption(speaker, text) {
  const normalizedSpeaker = String(speaker || '').trim().toLowerCase();
  let normalizedText = String(text || '').trim().toLowerCase();
  if (!normalizedText) return true;

  for (const pattern of UI_NOISE_PATTERNS) {
    normalizedText = normalizedText.replace(pattern, ' ').replace(/\s+/g, ' ').trim();
  }

  if (!normalizedText) return true;
  if (normalizedSpeaker === 'captions') return true;
  if (normalizedText === 'jump to bottom') return true;
  return false;
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
  if (isUiNoiseCaption(speaker, text)) return;
  if (alreadySent(speaker, text)) return;

  // Reject pure noise (no letters at all)
  if (/^[^a-zA-Z]*$/.test(text)) return;

  markSent(speaker, text);

  const caption = { speaker, text, timestamp: new Date().toISOString() };
  console.log('[Meet Capture] ✅', speaker + ':', text.slice(0, 80));

  chrome.runtime.sendMessage({ type: 'CAPTION_CAPTURED', caption }).catch(() => { });
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

  if (speaker) state.speaker = speaker;

  // No change since last check
  if (text === state.lastText) return;
  state.lastText = text;

  // Reset debounce timer
  if (state.timer) clearTimeout(state.timer);
  state.timer = setTimeout(() => {
    const inDom = document.contains(block);
    const finalText = inDom ? extractFromBlock(block).text : state.lastText;

    if (finalText && finalText !== state.lastSentText && finalText.length >= MIN_LENGTH) {
      sendCaption(state.speaker, finalText);
      state.lastSentText = finalText;
    }
    state.timer = null;
  }, DEBOUNCE_MS);
}

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
    if (m.addedNodes.length > 0) needsBlockScan = true;

    for (const node of m.removedNodes) {
      if (node.nodeType === Node.ELEMENT_NODE) {
        finalizeBlock(node);
        node.querySelectorAll && node.querySelectorAll('*').forEach(child => finalizeBlock(child));
      }
    }

    if (m.type === 'characterData' || m.type === 'childList') {
      let el = m.target.nodeType === Node.ELEMENT_NODE ? m.target : m.target.parentElement;
      while (el && el.parentElement !== container) {
        el = el.parentElement;
      }
      if (el && el.parentElement === container) {
        processBlock(el);
        needsBlockScan = false;
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

function pollBlocks() {
  if (!isCapturing) return;
  const container = getCaptionContainer();
  if (container) getCaptionBlocks(container).forEach(processBlock);
}

// ── Body-level observer: watch for caption container appearing/disappearing ─

function startDiscoveryObserver() {
  if (discoveryObserver) return;

  discoveryObserver = new MutationObserver(() => {
    if (!isCapturing) return;
    if (!mainObserver) {
      // The main observer isn't attached yet — try to find and attach
      const container = getCaptionContainer();
      if (container) {
        attachMainObserver(container);
      }
    }
  });

  discoveryObserver.observe(document.body, {
    childList: true,
    subtree: true,
  });

  // Also run periodic discovery in case MutationObserver misses it
  discoveryInterval = setInterval(() => {
    if (!isCapturing) return;
    if (!mainObserver) {
      const container = getCaptionContainer();
      if (container) {
        attachMainObserver(container);
      } else {
        console.log('[Meet Capture] ⏳ Still searching for caption container...');
      }
    } else {
      // Verify the container is still valid
      if (lastKnownContainer && !document.contains(lastKnownContainer)) {
        console.log('[Meet Capture] ⚠️ Caption container removed from DOM, re-discovering...');
        mainObserver.disconnect();
        mainObserver = null;
        lastKnownContainer = null;
      }
    }
  }, 3000);
}

function stopDiscoveryObserver() {
  if (discoveryObserver) { discoveryObserver.disconnect(); discoveryObserver = null; }
  if (discoveryInterval) { clearInterval(discoveryInterval); discoveryInterval = null; }
}

function attachMainObserver(container) {
  if (mainObserver) mainObserver.disconnect();

  mainObserver = new MutationObserver(handleMutations);
  mainObserver.observe(container, { childList: true, subtree: true, characterData: true });

  if (!pollInterval) {
    pollInterval = setInterval(pollBlocks, 800);
  }

  console.log('[Meet Capture] Started — container found:', container.tagName,
    (container.className || '').slice(0, 50), container.id || '');
  chrome.runtime.sendMessage({ type: 'CAPTURE_STARTED' }).catch(() => { });
}

// ── Start / Stop ───────────────────────────────────────────────────────────

function startCapture() {
  if (isCapturing) return;
  isCapturing = true;

  console.log('[Meet Capture] Starting capture...');

  const container = getCaptionContainer();
  if (container) {
    attachMainObserver(container);
  } else {
    console.warn('[Meet Capture] Caption container not found yet. Make sure captions (CC) are ON. Will keep searching...');
    chrome.runtime.sendMessage({ type: 'CAPTURE_STARTED' }).catch(() => { });
  }

  // Always start the discovery observer so we detect the container appearing
  startDiscoveryObserver();
}

function stopCapture() {
  isCapturing = false;

  if (mainObserver) { mainObserver.disconnect(); mainObserver = null; }
  if (pollInterval) { clearInterval(pollInterval); pollInterval = null; }
  stopDiscoveryObserver();
  lastKnownContainer = null;

  recentlySent.length = 0;
  console.log('[Meet Capture] Stopped');
  chrome.runtime.sendMessage({ type: 'CAPTURE_STOPPED' }).catch(() => { });
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

chrome.storage.onChanged.addListener((changes, areaName) => {
  if (areaName !== 'local' || !changes.currentSession) return;
  const nextSession = changes.currentSession.newValue;

  if (nextSession && !isCapturing) {
    console.log('[Meet Capture] Session detected in storage, starting capture');
    setTimeout(startCapture, 500);
  } else if (!nextSession && isCapturing) {
    console.log('[Meet Capture] Session cleared from storage, stopping capture');
    stopCapture();
  }
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
  }).catch(() => { });
}

// Auto-start if a session is active or if setting is enabled.
chrome.storage.local.get(['autoCapture', 'currentSession'], result => {
  if (result.currentSession) {
    setTimeout(startCapture, 1000);
    return;
  }
  if (result.autoCapture) setTimeout(startCapture, 3000);
});
