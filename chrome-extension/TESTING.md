# Testing Guide: Google Meet Note-Taker

This guide walks through testing the Chrome extension and backend integration.

## Pre-Test Checklist

- [ ] Backend server running (`npm start`)
- [ ] Chrome extension loaded (Developer mode)
- [ ] Google OAuth connected (Calendar + Docs scopes)
- [ ] OpenAI API key configured in `.env`
- [ ] Icon files added to `icons/` folder

## Test 1: Extension Installation

**Goal**: Verify extension loads without errors

1. Open `chrome://extensions`
2. Enable **Developer mode**
3. Click **Load unpacked** → Select `chrome-extension/` folder
4. Check for errors in extension card

**Expected**:
- ✅ Extension card shows "Google Meet Note-Taker"
- ✅ Version 1.0.0
- ✅ No errors displayed

**Debugging**:
- Errors? Click "Errors" button to see details
- Usually manifest.json syntax or missing files

---

## Test 2: Backend WebSocket Server

**Goal**: Verify WebSocket server is running

1. Start backend server:
   ```bash
   npm start
   ```

2. Check console output for:
   ```
   WebSocket server initialized on path: /meet-notes
   ```

3. Test WebSocket connection manually (optional):
   ```javascript
   // Open browser console on localhost:3000
   const ws = new WebSocket('ws://localhost:3000/meet-notes');
   ws.onopen = () => console.log('Connected!');
   ws.onerror = (err) => console.error('Error:', err);
   ```

**Expected**:
- ✅ Server shows "WebSocket server initialized"
- ✅ Test connection logs "Connected!"

**Debugging**:
- Connection refused? Check PORT in `.env` (default 3000)
- EADDRINUSE? Kill existing process: `netstat -ano | findstr :3000`

---

## Test 3: Content Script Injection

**Goal**: Verify content script loads on Meet pages

1. Join a Google Meet: [meet.google.com/new](https://meet.google.com/new)
2. Open DevTools (F12) → Console tab
3. Filter logs by: `[Meet Capture]`

**Expected**:
- ✅ Console shows: `[Meet Capture] Content script loaded`
- ✅ Console shows: `[Meet Capture] Meeting detected: abc-defg-hij`

**Debugging**:
- No logs? Check manifest.json `content_scripts.matches`
- Still no logs? Reload extension and refresh Meet tab

---

## Test 4: Caption Detection

**Goal**: Verify content script detects captions

1. In the Meet call, click **CC** button (bottom toolbar)
2. Speak or wait for auto-captions to appear
3. Open extension popup (click icon in toolbar)
4. Check DevTools console (filter: `[Meet Capture]`)

**Expected**:
- ✅ Console shows: `[Meet Capture] Found caption container: [selector]`
- ✅ Console shows: `[Meet Capture] Caption extracted: { speaker, text, timestamp }`

**Debugging**:
- "Caption container not found" → Captions not enabled in Meet
- Wait 3-5 seconds after enabling captions
- Try speaking louder or using test audio

---

## Test 5: Extension Popup UI

**Goal**: Verify popup displays meeting info

1. Click extension icon (while on Meet page)
2. Popup should open

**Expected**:
- ✅ Header: "Meet Note-Taker"
- ✅ Meeting info section visible
- ✅ "Start Capture" button enabled
- ✅ Footer shows: "Backend: Connected ✓"

**Debugging**:
- No meeting info? Check Calendar API connection
- "Backend: Disconnected"? WebSocket server not running
- Right-click icon → Inspect popup → Console for errors

---

## Test 6: Start Capture Session

**Goal**: Test full capture workflow

1. In the popup, click **Start Capture**
2. Observe UI changes

**Expected**:
- ✅ Button changes to "Stop & Save"
- ✅ Transcript container appears
- ✅ Status indicator shows "Capturing"
- ✅ Backend console logs: `[WebSocket] Session started: session_xyz`

**Debugging**:
- Error message in popup? Check popup console
- No backend log? WebSocket connection failed
- Check browser network tab for WebSocket connection

---

## Test 7: Live Caption Streaming

**Goal**: Verify captions appear in popup

1. With capture active, speak in the meeting
2. Watch the popup transcript section

**Expected**:
- ✅ Captions appear in popup within 1-2 seconds
- ✅ Speaker name displayed (or "Unknown")
- ✅ Caption count updates (e.g., "3 captions")
- ✅ Auto-scrolls to bottom

**Debugging**:
- No captions? Check Meet has CC enabled
- Content script console logs show extraction?
- Background script receiving messages?

---

## Test 8: Stop & Finalize

**Goal**: Test summary generation and doc creation

1. Click **Stop & Save** in popup
2. Wait for processing (3-10 seconds)

**Expected**:
- ✅ Loading spinner: "Saving notes..."
- ✅ Success screen appears
- ✅ Message: "Meeting notes created with X captions"
- ✅ "Open Google Doc" button visible
- ✅ Backend console logs:
  - `[Meet] Generating summary with OpenAI...`
  - `[Meet] Creating Google Doc...`
  - `[Meet] Shared with email@example.com`

**Debugging**:
- "Failed to create notes"? Check error message
- OpenAI API error? Check API key and quota
- Google Docs error? Verify Docs scope in OAuth
- Sharing failed? Check attendee emails in Calendar

---

## Test 9: Google Doc Verification

**Goal**: Verify created document is correct

1. Click **Open Google Doc** in success screen
2. Document should open in new tab

**Expected**:
- ✅ Doc title: "Meeting Name - Notes (MM/DD/YY)"
- ✅ Contains metadata section (date, attendees, duration)
- ✅ Contains AI Summary section with:
  - Key Discussion Points
  - Decisions Made
  - Action Items
  - Next Steps
- ✅ Contains Full Transcript section with all captions
- ✅ Shared with meeting attendees (check "Share" button)

**Debugging**:
- Doc empty? Check server logs for errors
- No summary? OpenAI request failed (check API key)
- Not shared? Attendees list empty in Calendar

---

## Test 10: Reconnection & Buffering

**Goal**: Test offline caption buffering

1. Start capture in a meeting
2. Stop backend server (Ctrl+C)
3. Speak in the meeting (captions should buffer locally)
4. Restart backend server
5. Captions should sync automatically

**Expected**:
- ✅ Popup shows "Backend: Disconnected"
- ✅ Captions still appear in popup (local capture)
- ✅ Background console logs: `Caption buffered (disconnected)`
- ✅ On reconnect: `Flushing X buffered captions`
- ✅ Popup shows "Backend: Connected ✓"

**Debugging**:
- Captions not buffering? Check background script logs
- Not reconnecting? Check exponential backoff logic
- Lost captions? Buffer cleared prematurely

---

## Test 11: Multiple Sessions

**Goal**: Test starting new session after completing one

1. Complete a capture session (Test 8)
2. Join a new Meet call
3. Start a new capture session

**Expected**:
- ✅ Previous session data cleared from popup
- ✅ New meeting info fetched
- ✅ New session ID generated
- ✅ Can capture and save second meeting independently

**Debugging**:
- Old captions appearing? Session not cleaned up
- Can't start new session? Check storage cleanup

---

## Test 12: Error Handling

**Goal**: Test various error scenarios

### A. No Calendar Event

1. Join a Meet without creating Calendar event
2. Click extension icon

**Expected**:
- ✅ Meeting info shows generic "Google Meet" title
- ✅ Can still capture and save notes
- ✅ Doc not auto-shared (no attendees)

### B. Captions Disabled

1. Join Meet without enabling captions
2. Click Start Capture

**Expected**:
- ✅ Content script logs: "Caption container not found"
- ✅ Retry after 2 seconds
- ✅ No captions in popup

### C. Backend Offline

1. Stop backend server
2. Click Start Capture

**Expected**:
- ✅ Popup shows "Backend: Disconnected"
- ✅ Start button disabled or shows error
- ✅ Cannot start new session

### D. OpenAI API Failure

1. Set invalid OpenAI API key
2. Complete a capture session
3. Click Stop & Save

**Expected**:
- ✅ Error message in popup
- ✅ Server logs OpenAI error
- ✅ Session not finalized

---

## Performance Tests

### Caption Volume

Test with high caption volume:

1. Start capture
2. Enable auto-captions in Meet
3. Play a video or speak continuously for 5 minutes

**Expected**:
- ✅ All captions captured (no drops)
- ✅ Popup remains responsive
- ✅ Memory usage stable (<100MB)

### Large Transcript

Test with 100+ captions:

1. Capture a 30-minute meeting
2. Generate summary and doc

**Expected**:
- ✅ Summary generated successfully
- ✅ Doc created (may take 5-10 seconds)
- ✅ Full transcript included in doc

---

## Browser Compatibility

Test on multiple browsers:

- [ ] Google Chrome (recommended)
- [ ] Microsoft Edge
- [ ] Brave Browser
- [ ] Opera

**Expected**: Should work on all Chromium-based browsers

---

## Success Criteria Summary

All tests passing means:

✅ Extension loads without errors
✅ Content script detects and captures captions
✅ WebSocket connects to backend
✅ Live captions stream to popup
✅ OpenAI generates summaries
✅ Google Docs created with correct content
✅ Docs shared with meeting attendees
✅ Offline buffering and reconnection works
✅ Error scenarios handled gracefully

---

## Next Steps After Testing

1. **Fix any failing tests** using debugging steps above
2. **Add custom icons** to make extension look professional
3. **Test in real meetings** with colleagues
4. **Gather feedback** on UI/UX
5. **Consider publishing** to Chrome Web Store (optional)

## Getting Help

If you encounter issues:

1. Check browser console logs (content script, background, popup)
2. Check backend server logs
3. Review TESTING.md debugging sections
4. Open GitHub issue with:
   - Test number that failed
   - Error messages
   - Browser console logs
   - Backend server logs

---

**Happy Testing! 🎉**
