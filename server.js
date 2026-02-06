require('dotenv').config();
const express = require('express');
const cors = require('cors');
const { google } = require('googleapis');
const OpenAI = require('openai');
const { Octokit } = require('octokit');
const fs = require('fs');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors());
app.use(express.json());
app.use(express.static('public'));

// OpenAI setup
const openai = new OpenAI({
    apiKey: process.env.OPENAI_API_KEY
});

// Gmail + Calendar OAuth setup
const SCOPES = [
    'https://www.googleapis.com/auth/gmail.readonly',
    'https://www.googleapis.com/auth/gmail.send',
    'https://www.googleapis.com/auth/gmail.modify',
    'https://www.googleapis.com/auth/gmail.labels',
    'https://www.googleapis.com/auth/gmail.compose',
    'https://www.googleapis.com/auth/gmail.settings.basic',
    'https://www.googleapis.com/auth/calendar'
];
const CALENDAR_SCOPE = 'https://www.googleapis.com/auth/calendar';

const TOKEN_PATH = path.join(process.env.USERPROFILE || process.env.HOME, '.gmail-mcp', 'token.json');
const GITHUB_TOKEN_PATH = path.join(process.env.USERPROFILE || process.env.HOME, '.gmail-mcp', 'github-token.json');

let oauth2Client = null;
let gmailClient = null;
let calendarClient = null;
let octokitClient = null;

function parseScopes(scopeString) {
    if (!scopeString || typeof scopeString !== 'string') return new Set();
    return new Set(scopeString.split(/\s+/).filter(Boolean));
}

function tokenHasScope(token, scope) {
    if (!token || !scope) return false;
    return parseScopes(token.scope).has(scope);
}

function readSavedToken() {
    try {
        if (!fs.existsSync(TOKEN_PATH)) return null;
        return JSON.parse(fs.readFileSync(TOKEN_PATH, 'utf8'));
    } catch (error) {
        console.error('Error reading saved Google token:', error.message);
        return null;
    }
}

function hasCalendarScope() {
    if (oauth2Client && tokenHasScope(oauth2Client.credentials, CALENDAR_SCOPE)) {
        return true;
    }
    const savedToken = readSavedToken();
    return tokenHasScope(savedToken, CALENDAR_SCOPE);
}

function getCalendarPermissionError(error) {
    const status = error?.code || error?.status || error?.response?.status;
    const message = String(error?.message || '');
    const looksLikeScopeError = /insufficient|permission|scope|forbidden|unauthorized/i.test(message);
    if ((status === 401 || status === 403) && looksLikeScopeError) {
        return 'Calendar permission is missing or expired. Please reconnect Google Calendar from the Calendar panel and try again.';
    }
    return null;
}

// Initialize OAuth client from environment variables
function initOAuthClient() {
    const clientId = process.env.GMAIL_CLIENT_ID || process.env.GOOGLE_CLIENT_ID;
    const clientSecret = process.env.GMAIL_CLIENT_SECRET || process.env.GOOGLE_CLIENT_SECRET;

    if (!clientId || !clientSecret || clientId === 'your_google_client_id_here') {
        console.log('  Google OAuth credentials not configured in .env');
        return false;
    }

    try {
        oauth2Client = new google.auth.OAuth2(
            clientId,
            clientSecret,
            `http://localhost:${PORT}/oauth2callback`
        );

        if (fs.existsSync(TOKEN_PATH)) {
            const token = JSON.parse(fs.readFileSync(TOKEN_PATH, 'utf8'));
            oauth2Client.setCredentials(token);
            gmailClient = google.gmail({ version: 'v1', auth: oauth2Client });
            calendarClient = tokenHasScope(token, CALENDAR_SCOPE)
                ? google.calendar({ version: 'v3', auth: oauth2Client })
                : null;
            if (calendarClient) {
                console.log('Gmail + Calendar clients initialized with existing token');
            } else {
                console.log('Gmail initialized. Calendar scope missing in token; reconnect required for Calendar tools.');
            }
        }
        return true;
    } catch (error) {
        console.error('Error initializing OAuth client:', error);
        return false;
    }
}

// Initialize GitHub client from saved token
function initGitHubClient() {
    try {
        if (fs.existsSync(GITHUB_TOKEN_PATH)) {
            const data = JSON.parse(fs.readFileSync(GITHUB_TOKEN_PATH, 'utf8'));
            if (data.token) {
                octokitClient = new Octokit({ auth: data.token });
                console.log('GitHub client initialized with saved token');
                return true;
            }
        }
    } catch (error) {
        console.error('Error initializing GitHub client:', error);
    }
    return false;
}

// ============================================================
//  HELPER UTILITIES
// ============================================================

function buildRawMessage({ to, subject, body, cc, bcc, inReplyTo, references, threadId }) {
    const lines = [];
    if (to) lines.push(`To: ${Array.isArray(to) ? to.join(', ') : to}`);
    if (cc) lines.push(`Cc: ${Array.isArray(cc) ? cc.join(', ') : cc}`);
    if (bcc) lines.push(`Bcc: ${Array.isArray(bcc) ? bcc.join(', ') : bcc}`);
    lines.push(`Subject: ${subject || ''}`);
    if (inReplyTo) lines.push(`In-Reply-To: ${inReplyTo}`);
    if (references) lines.push(`References: ${references}`);
    lines.push('Content-Type: text/html; charset=utf-8');
    lines.push('');
    lines.push(body || '');

    const raw = Buffer.from(lines.join('\r\n'))
        .toString('base64')
        .replace(/\+/g, '-')
        .replace(/\//g, '_')
        .replace(/=+$/, '');

    return raw;
}

function extractHeaders(headers, names) {
    const result = {};
    for (const name of names) {
        const h = headers.find(h => h.name.toLowerCase() === name.toLowerCase());
        result[name.toLowerCase()] = h ? h.value : '';
    }
    return result;
}

function extractBody(payload) {
    if (!payload) return '';
    if (payload.body && payload.body.data) {
        return Buffer.from(payload.body.data, 'base64').toString('utf8');
    }
    if (payload.parts) {
        // Prefer text/plain, fallback to text/html
        for (const mime of ['text/plain', 'text/html']) {
            for (const part of payload.parts) {
                if (part.mimeType === mime && part.body && part.body.data) {
                    return Buffer.from(part.body.data, 'base64').toString('utf8');
                }
                if (part.parts) {
                    const nested = extractBody(part);
                    if (nested) return nested;
                }
            }
        }
    }
    return '';
}

function extractAttachments(payload) {
    const attachments = [];
    function walk(parts) {
        if (!parts) return;
        for (const part of parts) {
            if (part.filename && part.filename.length > 0) {
                attachments.push({
                    filename: part.filename,
                    mimeType: part.mimeType,
                    size: part.body ? part.body.size : 0,
                    attachmentId: part.body ? part.body.attachmentId : null
                });
            }
            if (part.parts) walk(part.parts);
        }
    }
    if (payload.parts) walk(payload.parts);
    return attachments;
}

async function getEmailMetadata(messageId) {
    const email = await gmailClient.users.messages.get({
        userId: 'me',
        id: messageId,
        format: 'metadata',
        metadataHeaders: ['Subject', 'From', 'To', 'Date', 'Message-ID']
    });
    return email.data;
}

// ============================================================
//  25 GMAIL TOOL IMPLEMENTATIONS
// ============================================================

// 1. Send Email
async function sendEmail({ to, subject, body, cc, bcc }) {
    if (!gmailClient) throw new Error('Gmail not authenticated');
    const raw = buildRawMessage({ to, subject, body, cc, bcc });
    const result = await gmailClient.users.messages.send({
        userId: 'me',
        requestBody: { raw }
    });
    return { success: true, messageId: result.data.id, message: `Email sent to ${Array.isArray(to) ? to.join(', ') : to}` };
}

// 2. Search Emails
async function searchEmails({ query, maxResults = 20 }) {
    if (!gmailClient) throw new Error('Gmail not authenticated');
    const response = await gmailClient.users.messages.list({ userId: 'me', q: query, maxResults });
    if (!response.data.messages) return { results: [], totalEstimate: 0, message: 'No emails found' };

    const emails = await Promise.all(
        response.data.messages.slice(0, maxResults).map(async (msg) => {
            try {
                const data = await getEmailMetadata(msg.id);
                const h = extractHeaders(data.payload.headers, ['Subject', 'From', 'Date']);
                return { id: msg.id, snippet: data.snippet, subject: h.subject || '(no subject)', from: h.from, date: h.date, labelIds: data.labelIds };
            } catch { return null; }
        })
    );
    const filtered = emails.filter(Boolean);
    return { results: filtered, totalEstimate: response.data.resultSizeEstimate || filtered.length, message: `Found ${filtered.length} emails` };
}

// 3. Read Email (full content)
async function readEmail({ messageId }) {
    if (!gmailClient) throw new Error('Gmail not authenticated');
    const email = await gmailClient.users.messages.get({ userId: 'me', id: messageId, format: 'full' });
    const h = extractHeaders(email.data.payload.headers, ['Subject', 'From', 'To', 'Cc', 'Date', 'Message-ID']);
    const body = extractBody(email.data.payload);
    const attachments = extractAttachments(email.data.payload);
    return {
        id: messageId, threadId: email.data.threadId, labelIds: email.data.labelIds,
        snippet: email.data.snippet, subject: h.subject, from: h.from, to: h.to,
        cc: h.cc, date: h.date, messageIdHeader: h['message-id'],
        body, attachments, hasAttachments: attachments.length > 0
    };
}

// 4. List Emails
async function listEmails({ maxResults = 20, label = 'INBOX' }) {
    if (!gmailClient) throw new Error('Gmail not authenticated');
    const response = await gmailClient.users.messages.list({ userId: 'me', labelIds: [label], maxResults });
    if (!response.data.messages) return { results: [], message: 'No emails found' };

    const emails = await Promise.all(
        response.data.messages.map(async (msg) => {
            const data = await getEmailMetadata(msg.id);
            const h = extractHeaders(data.payload.headers, ['Subject', 'From', 'Date']);
            return { id: msg.id, subject: h.subject || '(no subject)', from: h.from, date: h.date, snippet: data.snippet, labelIds: data.labelIds };
        })
    );
    return { results: emails, message: `Listed ${emails.length} emails from ${label}` };
}

// 5. Trash Email
async function trashEmail({ messageId }) {
    if (!gmailClient) throw new Error('Gmail not authenticated');
    await gmailClient.users.messages.trash({ userId: 'me', id: messageId });
    return { success: true, message: `Email ${messageId} moved to trash` };
}

// 6. Modify Labels
async function modifyLabels({ messageId, addLabelIds = [], removeLabelIds = [] }) {
    if (!gmailClient) throw new Error('Gmail not authenticated');
    await gmailClient.users.messages.modify({
        userId: 'me', id: messageId,
        requestBody: { addLabelIds, removeLabelIds }
    });
    return { success: true, message: `Labels modified for email ${messageId}` };
}

// 7. Create Draft
async function createDraft({ to, subject, body, cc, bcc }) {
    if (!gmailClient) throw new Error('Gmail not authenticated');
    const raw = buildRawMessage({ to, subject, body, cc, bcc });
    const result = await gmailClient.users.drafts.create({
        userId: 'me',
        requestBody: { message: { raw } }
    });
    return { success: true, draftId: result.data.id, message: 'Draft created successfully' };
}

// 8. Reply to Email
async function replyToEmail({ messageId, body, cc }) {
    if (!gmailClient) throw new Error('Gmail not authenticated');
    const original = await gmailClient.users.messages.get({ userId: 'me', id: messageId, format: 'full' });
    const headers = original.data.payload.headers;
    const h = extractHeaders(headers, ['Subject', 'From', 'To', 'Message-ID']);

    const replyTo = h.from;
    const subject = h.subject.startsWith('Re:') ? h.subject : `Re: ${h.subject}`;
    const raw = buildRawMessage({
        to: replyTo, subject, body, cc,
        inReplyTo: h['message-id'],
        references: h['message-id']
    });

    const result = await gmailClient.users.messages.send({
        userId: 'me',
        requestBody: { raw, threadId: original.data.threadId }
    });
    return { success: true, messageId: result.data.id, message: `Reply sent to ${replyTo}` };
}

// 9. Forward Email
async function forwardEmail({ messageId, to, additionalMessage }) {
    if (!gmailClient) throw new Error('Gmail not authenticated');
    const original = await gmailClient.users.messages.get({ userId: 'me', id: messageId, format: 'full' });
    const h = extractHeaders(original.data.payload.headers, ['Subject', 'From', 'Date']);
    const originalBody = extractBody(original.data.payload);

    const forwardBody = `${additionalMessage || ''}\n\n---------- Forwarded message ----------\nFrom: ${h.from}\nDate: ${h.date}\nSubject: ${h.subject}\n\n${originalBody}`;
    const subject = h.subject.startsWith('Fwd:') ? h.subject : `Fwd: ${h.subject}`;

    const raw = buildRawMessage({ to, subject, body: forwardBody });
    const result = await gmailClient.users.messages.send({ userId: 'me', requestBody: { raw } });
    return { success: true, messageId: result.data.id, message: `Email forwarded to ${Array.isArray(to) ? to.join(', ') : to}` };
}

// 10. List Labels
async function listLabels() {
    if (!gmailClient) throw new Error('Gmail not authenticated');
    const response = await gmailClient.users.labels.list({ userId: 'me' });
    const labels = response.data.labels.map(l => ({
        id: l.id, name: l.name, type: l.type,
        messagesTotal: l.messagesTotal, messagesUnread: l.messagesUnread
    }));
    return { labels, message: `Found ${labels.length} labels` };
}

// 11. Create Label
async function createLabel({ name, backgroundColor, textColor }) {
    if (!gmailClient) throw new Error('Gmail not authenticated');
    const requestBody = {
        name,
        labelListVisibility: 'labelShow',
        messageListVisibility: 'show'
    };
    if (backgroundColor || textColor) {
        requestBody.color = {};
        if (backgroundColor) requestBody.color.backgroundColor = backgroundColor;
        if (textColor) requestBody.color.textColor = textColor;
    }
    const result = await gmailClient.users.labels.create({ userId: 'me', requestBody });
    return { success: true, labelId: result.data.id, name: result.data.name, message: `Label "${name}" created` };
}

// 12. Delete Label
async function deleteLabel({ labelId }) {
    if (!gmailClient) throw new Error('Gmail not authenticated');
    await gmailClient.users.labels.delete({ userId: 'me', id: labelId });
    return { success: true, message: `Label ${labelId} deleted` };
}

// 13. Mark as Read
async function markAsRead({ messageId }) {
    if (!gmailClient) throw new Error('Gmail not authenticated');
    await gmailClient.users.messages.modify({
        userId: 'me', id: messageId,
        requestBody: { removeLabelIds: ['UNREAD'] }
    });
    return { success: true, message: `Email ${messageId} marked as read` };
}

// 14. Mark as Unread
async function markAsUnread({ messageId }) {
    if (!gmailClient) throw new Error('Gmail not authenticated');
    await gmailClient.users.messages.modify({
        userId: 'me', id: messageId,
        requestBody: { addLabelIds: ['UNREAD'] }
    });
    return { success: true, message: `Email ${messageId} marked as unread` };
}

// 15. Star Email
async function starEmail({ messageId }) {
    if (!gmailClient) throw new Error('Gmail not authenticated');
    await gmailClient.users.messages.modify({
        userId: 'me', id: messageId,
        requestBody: { addLabelIds: ['STARRED'] }
    });
    return { success: true, message: `Email ${messageId} starred` };
}

// 16. Unstar Email
async function unstarEmail({ messageId }) {
    if (!gmailClient) throw new Error('Gmail not authenticated');
    await gmailClient.users.messages.modify({
        userId: 'me', id: messageId,
        requestBody: { removeLabelIds: ['STARRED'] }
    });
    return { success: true, message: `Email ${messageId} unstarred` };
}

// 17. Archive Email (remove from INBOX)
async function archiveEmail({ messageId }) {
    if (!gmailClient) throw new Error('Gmail not authenticated');
    await gmailClient.users.messages.modify({
        userId: 'me', id: messageId,
        requestBody: { removeLabelIds: ['INBOX'] }
    });
    return { success: true, message: `Email ${messageId} archived` };
}

// 18. Untrash Email
async function untrashEmail({ messageId }) {
    if (!gmailClient) throw new Error('Gmail not authenticated');
    await gmailClient.users.messages.untrash({ userId: 'me', id: messageId });
    return { success: true, message: `Email ${messageId} restored from trash` };
}

// 19. Get Thread (full conversation)
async function getThread({ threadId, maxMessages = 50 }) {
    if (!gmailClient) throw new Error('Gmail not authenticated');
    const thread = await gmailClient.users.threads.get({ userId: 'me', id: threadId, format: 'full' });
    const messages = thread.data.messages.slice(0, maxMessages).map(msg => {
        const h = extractHeaders(msg.payload.headers, ['Subject', 'From', 'To', 'Date']);
        return {
            id: msg.id, snippet: msg.snippet, subject: h.subject,
            from: h.from, to: h.to, date: h.date,
            body: extractBody(msg.payload), labelIds: msg.labelIds
        };
    });
    return { threadId, messageCount: messages.length, messages, message: `Thread has ${messages.length} messages` };
}

// 20. List Drafts
async function listDrafts({ maxResults = 20 }) {
    if (!gmailClient) throw new Error('Gmail not authenticated');
    const response = await gmailClient.users.drafts.list({ userId: 'me', maxResults });
    if (!response.data.drafts) return { drafts: [], message: 'No drafts found' };

    const drafts = await Promise.all(
        response.data.drafts.map(async (d) => {
            try {
                const draft = await gmailClient.users.drafts.get({ userId: 'me', id: d.id, format: 'metadata' });
                const h = extractHeaders(draft.data.message.payload.headers, ['Subject', 'To', 'Date']);
                return { draftId: d.id, messageId: draft.data.message.id, subject: h.subject || '(no subject)', to: h.to, date: h.date };
            } catch { return null; }
        })
    );
    const filtered = drafts.filter(Boolean);
    return { drafts: filtered, message: `Found ${filtered.length} drafts` };
}

// 21. Delete Draft
async function deleteDraft({ draftId }) {
    if (!gmailClient) throw new Error('Gmail not authenticated');
    await gmailClient.users.drafts.delete({ userId: 'me', id: draftId });
    return { success: true, message: `Draft ${draftId} deleted` };
}

// 22. Send Draft
async function sendDraft({ draftId }) {
    if (!gmailClient) throw new Error('Gmail not authenticated');
    const result = await gmailClient.users.drafts.send({ userId: 'me', requestBody: { id: draftId } });
    return { success: true, messageId: result.data.id, message: `Draft sent successfully` };
}

// 23. Get Attachment Info
async function getAttachmentInfo({ messageId }) {
    if (!gmailClient) throw new Error('Gmail not authenticated');
    const email = await gmailClient.users.messages.get({ userId: 'me', id: messageId, format: 'full' });
    const attachments = extractAttachments(email.data.payload);
    return { messageId, attachments, count: attachments.length, message: `Found ${attachments.length} attachment(s)` };
}

// 24. Get Profile
async function getProfile() {
    if (!gmailClient) throw new Error('Gmail not authenticated');
    const profile = await gmailClient.users.getProfile({ userId: 'me' });
    return {
        emailAddress: profile.data.emailAddress,
        messagesTotal: profile.data.messagesTotal,
        threadsTotal: profile.data.threadsTotal,
        historyId: profile.data.historyId,
        message: `Profile for ${profile.data.emailAddress}`
    };
}

// 25. Batch Modify Emails (batch mark read, trash, label, etc.)
async function batchModifyEmails({ messageIds, addLabelIds = [], removeLabelIds = [] }) {
    if (!gmailClient) throw new Error('Gmail not authenticated');
    if (!messageIds || messageIds.length === 0) throw new Error('No message IDs provided');

    await gmailClient.users.messages.batchModify({
        userId: 'me',
        requestBody: { ids: messageIds, addLabelIds, removeLabelIds }
    });
    return { success: true, count: messageIds.length, message: `Batch modified ${messageIds.length} emails` };
}

// ============================================================
//  15 GOOGLE CALENDAR TOOL IMPLEMENTATIONS
// ============================================================

// 1. List Events
async function listEvents({ calendarId = 'primary', maxResults = 10, timeMin, timeMax }) {
    if (!calendarClient) throw new Error('Calendar not authenticated');
    const params = { calendarId, maxResults, singleEvents: true, orderBy: 'startTime' };
    if (timeMin) params.timeMin = timeMin;
    else params.timeMin = new Date().toISOString();
    if (timeMax) params.timeMax = timeMax;
    const response = await calendarClient.events.list(params);
    const events = (response.data.items || []).map(e => ({
        id: e.id, summary: e.summary || '(no title)', status: e.status,
        start: e.start?.dateTime || e.start?.date, end: e.end?.dateTime || e.end?.date,
        location: e.location, description: e.description,
        attendees: (e.attendees || []).map(a => ({ email: a.email, responseStatus: a.responseStatus })),
        htmlLink: e.htmlLink
    }));
    return { events, message: `Found ${events.length} events` };
}

// 2. Get Event
async function getEvent({ calendarId = 'primary', eventId }) {
    if (!calendarClient) throw new Error('Calendar not authenticated');
    const e = (await calendarClient.events.get({ calendarId, eventId })).data;
    return {
        id: e.id, summary: e.summary, status: e.status, description: e.description,
        start: e.start?.dateTime || e.start?.date, end: e.end?.dateTime || e.end?.date,
        location: e.location, creator: e.creator, organizer: e.organizer,
        attendees: e.attendees || [], recurrence: e.recurrence, htmlLink: e.htmlLink,
        message: `Event: ${e.summary}`
    };
}

// 3. Create Event
async function createEvent({ calendarId = 'primary', summary, description, location, startDateTime, endDateTime, startDate, endDate, attendees, recurrence, timeZone }) {
    if (!calendarClient) throw new Error('Calendar not authenticated');
    const event = { summary };
    if (description) event.description = description;
    if (location) event.location = location;
    if (startDateTime) event.start = { dateTime: startDateTime, timeZone: timeZone || 'UTC' };
    else if (startDate) event.start = { date: startDate };
    if (endDateTime) event.end = { dateTime: endDateTime, timeZone: timeZone || 'UTC' };
    else if (endDate) event.end = { date: endDate };
    if (attendees) event.attendees = attendees.map(email => ({ email }));
    if (recurrence) event.recurrence = recurrence;
    const result = await calendarClient.events.insert({ calendarId, requestBody: event });
    return { success: true, eventId: result.data.id, htmlLink: result.data.htmlLink, message: `Event "${summary}" created` };
}

// 4. Update Event
async function updateEvent({ calendarId = 'primary', eventId, summary, description, location, startDateTime, endDateTime, startDate, endDate, timeZone }) {
    if (!calendarClient) throw new Error('Calendar not authenticated');
    const existing = (await calendarClient.events.get({ calendarId, eventId })).data;
    if (summary !== undefined) existing.summary = summary;
    if (description !== undefined) existing.description = description;
    if (location !== undefined) existing.location = location;
    if (startDateTime) existing.start = { dateTime: startDateTime, timeZone: timeZone || existing.start?.timeZone || 'UTC' };
    else if (startDate) existing.start = { date: startDate };
    if (endDateTime) existing.end = { dateTime: endDateTime, timeZone: timeZone || existing.end?.timeZone || 'UTC' };
    else if (endDate) existing.end = { date: endDate };
    const result = await calendarClient.events.update({ calendarId, eventId, requestBody: existing });
    return { success: true, eventId: result.data.id, message: `Event "${result.data.summary}" updated` };
}

// 5. Delete Event
async function deleteEvent({ calendarId = 'primary', eventId }) {
    if (!calendarClient) throw new Error('Calendar not authenticated');
    await calendarClient.events.delete({ calendarId, eventId });
    return { success: true, message: `Event ${eventId} deleted` };
}

// 6. List Calendars
async function listCalendars() {
    if (!calendarClient) throw new Error('Calendar not authenticated');
    const response = await calendarClient.calendarList.list();
    const calendars = (response.data.items || []).map(c => ({
        id: c.id, summary: c.summary, description: c.description,
        primary: c.primary || false, backgroundColor: c.backgroundColor, timeZone: c.timeZone
    }));
    return { calendars, message: `Found ${calendars.length} calendars` };
}

// 7. Create Calendar
async function createCalendar({ summary, description, timeZone }) {
    if (!calendarClient) throw new Error('Calendar not authenticated');
    const body = { summary };
    if (description) body.description = description;
    if (timeZone) body.timeZone = timeZone;
    const result = await calendarClient.calendars.insert({ requestBody: body });
    return { success: true, calendarId: result.data.id, message: `Calendar "${summary}" created` };
}

// 8. Quick Add Event
async function quickAddEvent({ calendarId = 'primary', text }) {
    if (!calendarClient) throw new Error('Calendar not authenticated');
    const result = await calendarClient.events.quickAdd({ calendarId, text });
    return { success: true, eventId: result.data.id, summary: result.data.summary, start: result.data.start?.dateTime || result.data.start?.date, message: `Quick event created: "${result.data.summary}"` };
}

// 9. Get Free/Busy
async function getFreeBusy({ timeMin, timeMax, calendarIds = ['primary'] }) {
    if (!calendarClient) throw new Error('Calendar not authenticated');
    const response = await calendarClient.freebusy.query({
        requestBody: {
            timeMin, timeMax,
            items: calendarIds.map(id => ({ id }))
        }
    });
    const calendars = {};
    for (const [id, data] of Object.entries(response.data.calendars || {})) {
        calendars[id] = { busy: data.busy || [] };
    }
    return { calendars, message: 'Free/busy info retrieved' };
}

// 10. List Recurring Event Instances
async function listRecurringInstances({ calendarId = 'primary', eventId, maxResults = 10, timeMin, timeMax }) {
    if (!calendarClient) throw new Error('Calendar not authenticated');
    const params = { calendarId, eventId, maxResults };
    if (timeMin) params.timeMin = timeMin;
    if (timeMax) params.timeMax = timeMax;
    const response = await calendarClient.events.instances(params);
    const instances = (response.data.items || []).map(e => ({
        id: e.id, summary: e.summary,
        start: e.start?.dateTime || e.start?.date, end: e.end?.dateTime || e.end?.date,
        status: e.status
    }));
    return { instances, message: `Found ${instances.length} instances` };
}

// 11. Move Event
async function moveEvent({ calendarId = 'primary', eventId, destinationCalendarId }) {
    if (!calendarClient) throw new Error('Calendar not authenticated');
    const result = await calendarClient.events.move({ calendarId, eventId, destination: destinationCalendarId });
    return { success: true, eventId: result.data.id, message: `Event moved to calendar ${destinationCalendarId}` };
}

// 12. Update Event Attendees
async function updateEventAttendees({ calendarId = 'primary', eventId, addAttendees = [], removeAttendees = [] }) {
    if (!calendarClient) throw new Error('Calendar not authenticated');
    const existing = (await calendarClient.events.get({ calendarId, eventId })).data;
    let attendees = existing.attendees || [];
    if (addAttendees.length > 0) {
        const existingEmails = new Set(attendees.map(a => a.email));
        for (const email of addAttendees) {
            if (!existingEmails.has(email)) attendees.push({ email });
        }
    }
    if (removeAttendees.length > 0) {
        const removeSet = new Set(removeAttendees);
        attendees = attendees.filter(a => !removeSet.has(a.email));
    }
    existing.attendees = attendees;
    const result = await calendarClient.events.update({ calendarId, eventId, requestBody: existing });
    return { success: true, attendeeCount: (result.data.attendees || []).length, message: `Attendees updated for event "${result.data.summary}"` };
}

// 13. Get Calendar Colors
async function getCalendarColors() {
    if (!calendarClient) throw new Error('Calendar not authenticated');
    const response = await calendarClient.colors.get();
    return { calendar: response.data.calendar, event: response.data.event, message: 'Calendar colors retrieved' };
}

// 14. Clear Calendar
async function clearCalendar({ calendarId }) {
    if (!calendarClient) throw new Error('Calendar not authenticated');
    await calendarClient.calendars.clear({ calendarId });
    return { success: true, message: `Calendar ${calendarId} cleared` };
}

// 15. Watch Events
async function watchEvents({ calendarId = 'primary', webhookUrl }) {
    if (!calendarClient) throw new Error('Calendar not authenticated');
    const result = await calendarClient.events.watch({
        calendarId,
        requestBody: {
            id: `watch-${Date.now()}`,
            type: 'web_hook',
            address: webhookUrl
        }
    });
    return { success: true, channelId: result.data.id, expiration: result.data.expiration, message: 'Watch channel created' };
}

// ============================================================
//  20 GITHUB TOOL IMPLEMENTATIONS
// ============================================================

// 1. List Repos
async function listRepos({ username, sort = 'updated', perPage = 30 }) {
    if (!octokitClient) throw new Error('GitHub not connected');
    let response;
    if (username) {
        response = await octokitClient.rest.repos.listForUser({ username, sort, per_page: perPage });
    } else {
        response = await octokitClient.rest.repos.listForAuthenticatedUser({ sort, per_page: perPage });
    }
    const repos = response.data.map(r => ({
        name: r.name, full_name: r.full_name, description: r.description,
        private: r.private, language: r.language, stars: r.stargazers_count,
        forks: r.forks_count, open_issues: r.open_issues_count,
        url: r.html_url, updated_at: r.updated_at
    }));
    return { repos, message: `Found ${repos.length} repositories` };
}

// 2. Get Repo
async function getRepo({ owner, repo }) {
    if (!octokitClient) throw new Error('GitHub not connected');
    const r = (await octokitClient.rest.repos.get({ owner, repo })).data;
    return {
        name: r.name, full_name: r.full_name, description: r.description,
        private: r.private, language: r.language, stars: r.stargazers_count,
        forks: r.forks_count, open_issues: r.open_issues_count,
        default_branch: r.default_branch, created_at: r.created_at,
        updated_at: r.updated_at, url: r.html_url, topics: r.topics,
        message: `Repository: ${r.full_name}`
    };
}

// 3. Create Repo
async function createRepo({ name, description, isPrivate = false, autoInit = true }) {
    if (!octokitClient) throw new Error('GitHub not connected');
    const r = (await octokitClient.rest.repos.createForAuthenticatedUser({
        name, description, private: isPrivate, auto_init: autoInit
    })).data;
    return { success: true, name: r.name, full_name: r.full_name, url: r.html_url, message: `Repository "${r.full_name}" created` };
}

// 4. List Issues
async function listIssues({ owner, repo, state = 'open', labels, perPage = 30 }) {
    if (!octokitClient) throw new Error('GitHub not connected');
    const params = { owner, repo, state, per_page: perPage };
    if (labels) params.labels = labels;
    const response = await octokitClient.rest.issues.listForRepo(params);
    const issues = response.data.filter(i => !i.pull_request).map(i => ({
        number: i.number, title: i.title, state: i.state,
        user: i.user?.login, labels: i.labels.map(l => l.name),
        created_at: i.created_at, updated_at: i.updated_at,
        comments: i.comments, url: i.html_url
    }));
    return { issues, message: `Found ${issues.length} issues` };
}

// 5. Create Issue
async function createIssue({ owner, repo, title, body, labels, assignees }) {
    if (!octokitClient) throw new Error('GitHub not connected');
    const params = { owner, repo, title };
    if (body) params.body = body;
    if (labels) params.labels = labels;
    if (assignees) params.assignees = assignees;
    const i = (await octokitClient.rest.issues.create(params)).data;
    return { success: true, number: i.number, title: i.title, url: i.html_url, message: `Issue #${i.number} created: "${i.title}"` };
}

// 6. Update Issue
async function updateIssue({ owner, repo, issueNumber, title, body, state, labels, assignees }) {
    if (!octokitClient) throw new Error('GitHub not connected');
    const params = { owner, repo, issue_number: issueNumber };
    if (title !== undefined) params.title = title;
    if (body !== undefined) params.body = body;
    if (state) params.state = state;
    if (labels) params.labels = labels;
    if (assignees) params.assignees = assignees;
    const i = (await octokitClient.rest.issues.update(params)).data;
    return { success: true, number: i.number, title: i.title, state: i.state, url: i.html_url, message: `Issue #${i.number} updated` };
}

// 7. List Pull Requests
async function listPullRequests({ owner, repo, state = 'open', perPage = 30 }) {
    if (!octokitClient) throw new Error('GitHub not connected');
    const response = await octokitClient.rest.pulls.list({ owner, repo, state, per_page: perPage });
    const prs = response.data.map(p => ({
        number: p.number, title: p.title, state: p.state,
        user: p.user?.login, head: p.head?.ref, base: p.base?.ref,
        created_at: p.created_at, updated_at: p.updated_at,
        mergeable: p.mergeable, draft: p.draft, url: p.html_url
    }));
    return { pullRequests: prs, message: `Found ${prs.length} pull requests` };
}

// 8. Get Pull Request
async function getPullRequest({ owner, repo, pullNumber }) {
    if (!octokitClient) throw new Error('GitHub not connected');
    const p = (await octokitClient.rest.pulls.get({ owner, repo, pull_number: pullNumber })).data;
    return {
        number: p.number, title: p.title, state: p.state, body: p.body,
        user: p.user?.login, head: p.head?.ref, base: p.base?.ref,
        mergeable: p.mergeable, merged: p.merged, draft: p.draft,
        additions: p.additions, deletions: p.deletions, changed_files: p.changed_files,
        created_at: p.created_at, url: p.html_url,
        message: `PR #${p.number}: ${p.title}`
    };
}

// 9. Create Pull Request
async function createPullRequest({ owner, repo, title, body, head, base }) {
    if (!octokitClient) throw new Error('GitHub not connected');
    const p = (await octokitClient.rest.pulls.create({ owner, repo, title, body, head, base })).data;
    return { success: true, number: p.number, title: p.title, url: p.html_url, message: `PR #${p.number} created: "${p.title}"` };
}

// 10. Merge Pull Request
async function mergePullRequest({ owner, repo, pullNumber, mergeMethod = 'merge', commitMessage }) {
    if (!octokitClient) throw new Error('GitHub not connected');
    const params = { owner, repo, pull_number: pullNumber, merge_method: mergeMethod };
    if (commitMessage) params.commit_message = commitMessage;
    const result = (await octokitClient.rest.pulls.merge(params)).data;
    return { success: true, merged: result.merged, sha: result.sha, message: result.message };
}

// 11. List Branches
async function listBranches({ owner, repo, perPage = 30 }) {
    if (!octokitClient) throw new Error('GitHub not connected');
    const response = await octokitClient.rest.repos.listBranches({ owner, repo, per_page: perPage });
    const branches = response.data.map(b => ({
        name: b.name, sha: b.commit.sha, protected: b.protected
    }));
    return { branches, message: `Found ${branches.length} branches` };
}

// 12. Create Branch
async function createBranch({ owner, repo, branchName, fromBranch = 'main' }) {
    if (!octokitClient) throw new Error('GitHub not connected');
    const ref = await octokitClient.rest.git.getRef({ owner, repo, ref: `heads/${fromBranch}` });
    const sha = ref.data.object.sha;
    await octokitClient.rest.git.createRef({ owner, repo, ref: `refs/heads/${branchName}`, sha });
    return { success: true, branch: branchName, fromSha: sha, message: `Branch "${branchName}" created from "${fromBranch}"` };
}

// 13. Get File Content
async function getFileContent({ owner, repo, filePath, ref }) {
    if (!octokitClient) throw new Error('GitHub not connected');
    const params = { owner, repo, path: filePath };
    if (ref) params.ref = ref;
    const response = await octokitClient.rest.repos.getContent(params);
    const data = response.data;
    let content = '';
    if (data.content) {
        content = Buffer.from(data.content, 'base64').toString('utf8');
    }
    return { name: data.name, path: data.path, sha: data.sha, size: data.size, content, url: data.html_url, message: `File: ${data.path}` };
}

// 14. Create or Update File
async function createOrUpdateFile({ owner, repo, filePath, content, message, branch, sha }) {
    if (!octokitClient) throw new Error('GitHub not connected');
    const params = {
        owner, repo, path: filePath,
        message, content: Buffer.from(content).toString('base64')
    };
    if (branch) params.branch = branch;
    if (sha) params.sha = sha;
    const result = (await octokitClient.rest.repos.createOrUpdateFileContents(params)).data;
    return { success: true, path: result.content.path, sha: result.content.sha, url: result.content.html_url, message: `File "${filePath}" ${sha ? 'updated' : 'created'}` };
}

// 15. Search Repos
async function searchRepos({ query, sort = 'stars', perPage = 20 }) {
    if (!octokitClient) throw new Error('GitHub not connected');
    const response = await octokitClient.rest.search.repos({ q: query, sort, per_page: perPage });
    const repos = response.data.items.map(r => ({
        name: r.name, full_name: r.full_name, description: r.description,
        language: r.language, stars: r.stargazers_count, forks: r.forks_count,
        url: r.html_url
    }));
    return { total_count: response.data.total_count, repos, message: `Found ${response.data.total_count} repos matching "${query}"` };
}

// 16. Search Code
async function searchCode({ query, perPage = 20 }) {
    if (!octokitClient) throw new Error('GitHub not connected');
    const response = await octokitClient.rest.search.code({ q: query, per_page: perPage });
    const results = response.data.items.map(item => ({
        name: item.name, path: item.path,
        repository: item.repository.full_name,
        url: item.html_url
    }));
    return { total_count: response.data.total_count, results, message: `Found ${response.data.total_count} code results` };
}

// 17. List Commits
async function listCommits({ owner, repo, sha, perPage = 20 }) {
    if (!octokitClient) throw new Error('GitHub not connected');
    const params = { owner, repo, per_page: perPage };
    if (sha) params.sha = sha;
    const response = await octokitClient.rest.repos.listCommits(params);
    const commits = response.data.map(c => ({
        sha: c.sha.slice(0, 7), message: c.commit.message,
        author: c.commit.author?.name, date: c.commit.author?.date,
        url: c.html_url
    }));
    return { commits, message: `Found ${commits.length} commits` };
}

// 18. Get User Profile
async function getUserProfile({ username }) {
    if (!octokitClient) throw new Error('GitHub not connected');
    let u;
    if (username) {
        u = (await octokitClient.rest.users.getByUsername({ username })).data;
    } else {
        u = (await octokitClient.rest.users.getAuthenticated()).data;
    }
    return {
        login: u.login, name: u.name, bio: u.bio, company: u.company,
        location: u.location, public_repos: u.public_repos,
        followers: u.followers, following: u.following,
        avatar_url: u.avatar_url, url: u.html_url,
        message: `GitHub user: ${u.login}`
    };
}

// 19. List Notifications
async function listNotifications({ all = false, perPage = 20 }) {
    if (!octokitClient) throw new Error('GitHub not connected');
    const response = await octokitClient.rest.activity.listNotificationsForAuthenticatedUser({ all, per_page: perPage });
    const notifications = response.data.map(n => ({
        id: n.id, reason: n.reason, unread: n.unread,
        subject: { title: n.subject.title, type: n.subject.type },
        repository: n.repository.full_name,
        updated_at: n.updated_at
    }));
    return { notifications, message: `Found ${notifications.length} notifications` };
}

// 20. List Gists
async function listGists({ perPage = 20 }) {
    if (!octokitClient) throw new Error('GitHub not connected');
    const response = await octokitClient.rest.gists.list({ per_page: perPage });
    const gists = response.data.map(g => ({
        id: g.id, description: g.description,
        files: Object.keys(g.files),
        public: g.public, created_at: g.created_at,
        url: g.html_url
    }));
    return { gists, message: `Found ${gists.length} gists` };
}

// ============================================================
//  TOOL DEFINITIONS FOR OPENAI (25 Gmail + 15 Calendar + 20 GitHub = 60 Tools)
// ============================================================
const gmailTools = [
    {
        type: "function",
        function: {
            name: "send_email",
            description: "Send a new email to recipients. Supports CC/BCC. Body can be plain text or HTML.",
            parameters: {
                type: "object",
                properties: {
                    to: { type: "array", items: { type: "string" }, description: "Recipient email addresses" },
                    subject: { type: "string", description: "Email subject line" },
                    body: { type: "string", description: "Email body (plain text or HTML)" },
                    cc: { type: "array", items: { type: "string" }, description: "CC addresses" },
                    bcc: { type: "array", items: { type: "string" }, description: "BCC addresses" }
                },
                required: ["to", "subject", "body"]
            }
        }
    },
    {
        type: "function",
        function: {
            name: "search_emails",
            description: "Search emails using Gmail query syntax. Examples: 'from:boss', 'is:unread', 'subject:invoice', 'has:attachment', 'newer_than:2d', 'label:important'",
            parameters: {
                type: "object",
                properties: {
                    query: { type: "string", description: "Gmail search query" },
                    maxResults: { type: "integer", description: "Maximum results to return (default 20, max 100)" }
                },
                required: ["query"]
            }
        }
    },
    {
        type: "function",
        function: {
            name: "read_email",
            description: "Read the full content of an email including body, headers, and attachment info. Use after searching to get details.",
            parameters: {
                type: "object",
                properties: {
                    messageId: { type: "string", description: "The Gmail message ID" }
                },
                required: ["messageId"]
            }
        }
    },
    {
        type: "function",
        function: {
            name: "list_emails",
            description: "List recent emails from a specific label/folder. Common labels: INBOX, SENT, DRAFT, TRASH, SPAM, STARRED, IMPORTANT, UNREAD",
            parameters: {
                type: "object",
                properties: {
                    maxResults: { type: "integer", description: "Number of emails to return (default 20)" },
                    label: { type: "string", description: "Gmail label ID (default: INBOX)" }
                }
            }
        }
    },
    {
        type: "function",
        function: {
            name: "trash_email",
            description: "Move an email to the Trash folder",
            parameters: {
                type: "object",
                properties: {
                    messageId: { type: "string", description: "The Gmail message ID to trash" }
                },
                required: ["messageId"]
            }
        }
    },
    {
        type: "function",
        function: {
            name: "modify_labels",
            description: "Add or remove labels from an email. Use for custom label management.",
            parameters: {
                type: "object",
                properties: {
                    messageId: { type: "string", description: "The Gmail message ID" },
                    addLabelIds: { type: "array", items: { type: "string" }, description: "Label IDs to add" },
                    removeLabelIds: { type: "array", items: { type: "string" }, description: "Label IDs to remove" }
                },
                required: ["messageId"]
            }
        }
    },
    {
        type: "function",
        function: {
            name: "create_draft",
            description: "Create a new email draft (saved but not sent). Useful for composing emails to review before sending.",
            parameters: {
                type: "object",
                properties: {
                    to: { type: "array", items: { type: "string" }, description: "Recipients" },
                    subject: { type: "string", description: "Subject line" },
                    body: { type: "string", description: "Email body" },
                    cc: { type: "array", items: { type: "string" }, description: "CC addresses" },
                    bcc: { type: "array", items: { type: "string" }, description: "BCC addresses" }
                },
                required: ["subject", "body"]
            }
        }
    },
    {
        type: "function",
        function: {
            name: "reply_to_email",
            description: "Reply to an existing email. Automatically sets the correct thread, subject, and reply-to headers.",
            parameters: {
                type: "object",
                properties: {
                    messageId: { type: "string", description: "The message ID to reply to" },
                    body: { type: "string", description: "Reply body text" },
                    cc: { type: "array", items: { type: "string" }, description: "Additional CC addresses" }
                },
                required: ["messageId", "body"]
            }
        }
    },
    {
        type: "function",
        function: {
            name: "forward_email",
            description: "Forward an existing email to new recipients with an optional additional message.",
            parameters: {
                type: "object",
                properties: {
                    messageId: { type: "string", description: "The message ID to forward" },
                    to: { type: "array", items: { type: "string" }, description: "Forward recipients" },
                    additionalMessage: { type: "string", description: "Optional message to add above the forwarded content" }
                },
                required: ["messageId", "to"]
            }
        }
    },
    {
        type: "function",
        function: {
            name: "list_labels",
            description: "List all Gmail labels/folders including system labels (INBOX, SENT, etc.) and custom labels.",
            parameters: { type: "object", properties: {} }
        }
    },
    {
        type: "function",
        function: {
            name: "create_label",
            description: "Create a new Gmail label/folder for organizing emails.",
            parameters: {
                type: "object",
                properties: {
                    name: { type: "string", description: "Label name (e.g. 'Projects/ClientA')" },
                    backgroundColor: { type: "string", description: "Background color hex (e.g. '#16a765')" },
                    textColor: { type: "string", description: "Text color hex (e.g. '#ffffff')" }
                },
                required: ["name"]
            }
        }
    },
    {
        type: "function",
        function: {
            name: "delete_label",
            description: "Delete a Gmail label. Only works on user-created labels, not system labels.",
            parameters: {
                type: "object",
                properties: {
                    labelId: { type: "string", description: "The label ID to delete" }
                },
                required: ["labelId"]
            }
        }
    },
    {
        type: "function",
        function: {
            name: "mark_as_read",
            description: "Mark a specific email as read by removing the UNREAD label.",
            parameters: {
                type: "object",
                properties: {
                    messageId: { type: "string", description: "The Gmail message ID" }
                },
                required: ["messageId"]
            }
        }
    },
    {
        type: "function",
        function: {
            name: "mark_as_unread",
            description: "Mark a specific email as unread by adding the UNREAD label.",
            parameters: {
                type: "object",
                properties: {
                    messageId: { type: "string", description: "The Gmail message ID" }
                },
                required: ["messageId"]
            }
        }
    },
    {
        type: "function",
        function: {
            name: "star_email",
            description: "Star/flag an important email for quick access later.",
            parameters: {
                type: "object",
                properties: {
                    messageId: { type: "string", description: "The Gmail message ID to star" }
                },
                required: ["messageId"]
            }
        }
    },
    {
        type: "function",
        function: {
            name: "unstar_email",
            description: "Remove the star from an email.",
            parameters: {
                type: "object",
                properties: {
                    messageId: { type: "string", description: "The Gmail message ID to unstar" }
                },
                required: ["messageId"]
            }
        }
    },
    {
        type: "function",
        function: {
            name: "archive_email",
            description: "Archive an email by removing it from the Inbox. The email is still searchable but won't appear in Inbox.",
            parameters: {
                type: "object",
                properties: {
                    messageId: { type: "string", description: "The Gmail message ID to archive" }
                },
                required: ["messageId"]
            }
        }
    },
    {
        type: "function",
        function: {
            name: "untrash_email",
            description: "Restore an email from Trash back to the Inbox.",
            parameters: {
                type: "object",
                properties: {
                    messageId: { type: "string", description: "The Gmail message ID to restore" }
                },
                required: ["messageId"]
            }
        }
    },
    {
        type: "function",
        function: {
            name: "get_thread",
            description: "Get a full email thread/conversation. Returns all messages in the thread in chronological order.",
            parameters: {
                type: "object",
                properties: {
                    threadId: { type: "string", description: "The Gmail thread ID" },
                    maxMessages: { type: "integer", description: "Max messages to return from thread (default 50)" }
                },
                required: ["threadId"]
            }
        }
    },
    {
        type: "function",
        function: {
            name: "list_drafts",
            description: "List all saved email drafts.",
            parameters: {
                type: "object",
                properties: {
                    maxResults: { type: "integer", description: "Max drafts to return (default 20)" }
                }
            }
        }
    },
    {
        type: "function",
        function: {
            name: "delete_draft",
            description: "Permanently delete a draft email.",
            parameters: {
                type: "object",
                properties: {
                    draftId: { type: "string", description: "The draft ID to delete" }
                },
                required: ["draftId"]
            }
        }
    },
    {
        type: "function",
        function: {
            name: "send_draft",
            description: "Send an existing draft email immediately.",
            parameters: {
                type: "object",
                properties: {
                    draftId: { type: "string", description: "The draft ID to send" }
                },
                required: ["draftId"]
            }
        }
    },
    {
        type: "function",
        function: {
            name: "get_attachment_info",
            description: "Get information about attachments in an email (filenames, sizes, types).",
            parameters: {
                type: "object",
                properties: {
                    messageId: { type: "string", description: "The Gmail message ID" }
                },
                required: ["messageId"]
            }
        }
    },
    {
        type: "function",
        function: {
            name: "get_profile",
            description: "Get the authenticated Gmail user's profile info including email address, total messages, and total threads.",
            parameters: { type: "object", properties: {} }
        }
    },
    {
        type: "function",
        function: {
            name: "batch_modify_emails",
            description: "Apply label changes to multiple emails at once. Useful for bulk operations like 'mark all as read', 'archive all from sender', etc.",
            parameters: {
                type: "object",
                properties: {
                    messageIds: { type: "array", items: { type: "string" }, description: "Array of message IDs to modify" },
                    addLabelIds: { type: "array", items: { type: "string" }, description: "Labels to add to all messages" },
                    removeLabelIds: { type: "array", items: { type: "string" }, description: "Labels to remove from all messages" }
                },
                required: ["messageIds"]
            }
        }
    }
];

const calendarTools = [
    { type: "function", function: { name: "list_events", description: "List upcoming calendar events. Can filter by time range.", parameters: { type: "object", properties: { calendarId: { type: "string", description: "Calendar ID (default: primary)" }, maxResults: { type: "integer", description: "Max events to return (default 10)" }, timeMin: { type: "string", description: "Start time in ISO 8601 format" }, timeMax: { type: "string", description: "End time in ISO 8601 format" } } } } },
    { type: "function", function: { name: "get_event", description: "Get full details of a specific calendar event.", parameters: { type: "object", properties: { calendarId: { type: "string", description: "Calendar ID (default: primary)" }, eventId: { type: "string", description: "The event ID" } }, required: ["eventId"] } } },
    { type: "function", function: { name: "create_event", description: "Create a new calendar event with optional attendees, location, and recurrence.", parameters: { type: "object", properties: { calendarId: { type: "string", description: "Calendar ID (default: primary)" }, summary: { type: "string", description: "Event title" }, description: { type: "string", description: "Event description" }, location: { type: "string", description: "Event location" }, startDateTime: { type: "string", description: "Start datetime in ISO 8601 (for timed events)" }, endDateTime: { type: "string", description: "End datetime in ISO 8601 (for timed events)" }, startDate: { type: "string", description: "Start date YYYY-MM-DD (for all-day events)" }, endDate: { type: "string", description: "End date YYYY-MM-DD (for all-day events)" }, attendees: { type: "array", items: { type: "string" }, description: "Attendee email addresses" }, recurrence: { type: "array", items: { type: "string" }, description: "RRULE strings, e.g. ['RRULE:FREQ=WEEKLY;COUNT=5']" }, timeZone: { type: "string", description: "Time zone (default: UTC)" } }, required: ["summary"] } } },
    { type: "function", function: { name: "update_event", description: "Update an existing calendar event's title, time, location, or description.", parameters: { type: "object", properties: { calendarId: { type: "string", description: "Calendar ID (default: primary)" }, eventId: { type: "string", description: "The event ID to update" }, summary: { type: "string", description: "New event title" }, description: { type: "string", description: "New description" }, location: { type: "string", description: "New location" }, startDateTime: { type: "string", description: "New start datetime" }, endDateTime: { type: "string", description: "New end datetime" }, startDate: { type: "string", description: "New start date (all-day)" }, endDate: { type: "string", description: "New end date (all-day)" }, timeZone: { type: "string", description: "Time zone" } }, required: ["eventId"] } } },
    { type: "function", function: { name: "delete_event", description: "Delete a calendar event.", parameters: { type: "object", properties: { calendarId: { type: "string", description: "Calendar ID (default: primary)" }, eventId: { type: "string", description: "The event ID to delete" } }, required: ["eventId"] } } },
    { type: "function", function: { name: "list_calendars", description: "List all calendars accessible to the user.", parameters: { type: "object", properties: {} } } },
    { type: "function", function: { name: "create_calendar", description: "Create a new calendar.", parameters: { type: "object", properties: { summary: { type: "string", description: "Calendar name" }, description: { type: "string", description: "Calendar description" }, timeZone: { type: "string", description: "Time zone" } }, required: ["summary"] } } },
    { type: "function", function: { name: "quick_add_event", description: "Quickly create an event using natural language (e.g. 'Meeting tomorrow at 3pm').", parameters: { type: "object", properties: { calendarId: { type: "string", description: "Calendar ID (default: primary)" }, text: { type: "string", description: "Natural language event description" } }, required: ["text"] } } },
    { type: "function", function: { name: "get_free_busy", description: "Check free/busy status for calendars in a time range.", parameters: { type: "object", properties: { timeMin: { type: "string", description: "Start of time range (ISO 8601)" }, timeMax: { type: "string", description: "End of time range (ISO 8601)" }, calendarIds: { type: "array", items: { type: "string" }, description: "Calendar IDs to check (default: ['primary'])" } }, required: ["timeMin", "timeMax"] } } },
    { type: "function", function: { name: "list_recurring_instances", description: "List individual occurrences of a recurring event.", parameters: { type: "object", properties: { calendarId: { type: "string", description: "Calendar ID (default: primary)" }, eventId: { type: "string", description: "The recurring event ID" }, maxResults: { type: "integer", description: "Max instances to return" }, timeMin: { type: "string", description: "Start time filter" }, timeMax: { type: "string", description: "End time filter" } }, required: ["eventId"] } } },
    { type: "function", function: { name: "move_event", description: "Move an event from one calendar to another.", parameters: { type: "object", properties: { calendarId: { type: "string", description: "Source calendar ID (default: primary)" }, eventId: { type: "string", description: "The event ID to move" }, destinationCalendarId: { type: "string", description: "Destination calendar ID" } }, required: ["eventId", "destinationCalendarId"] } } },
    { type: "function", function: { name: "update_event_attendees", description: "Add or remove attendees from a calendar event.", parameters: { type: "object", properties: { calendarId: { type: "string", description: "Calendar ID (default: primary)" }, eventId: { type: "string", description: "The event ID" }, addAttendees: { type: "array", items: { type: "string" }, description: "Email addresses to add" }, removeAttendees: { type: "array", items: { type: "string" }, description: "Email addresses to remove" } }, required: ["eventId"] } } },
    { type: "function", function: { name: "get_calendar_colors", description: "Get available color options for calendars and events.", parameters: { type: "object", properties: {} } } },
    { type: "function", function: { name: "clear_calendar", description: "Clear all events from a calendar. WARNING: This is destructive!", parameters: { type: "object", properties: { calendarId: { type: "string", description: "The calendar ID to clear (cannot be primary)" } }, required: ["calendarId"] } } },
    { type: "function", function: { name: "watch_events", description: "Set up push notifications for calendar changes (requires a public webhook URL).", parameters: { type: "object", properties: { calendarId: { type: "string", description: "Calendar ID (default: primary)" }, webhookUrl: { type: "string", description: "Public webhook URL to receive notifications" } }, required: ["webhookUrl"] } } }
];

const githubTools = [
    { type: "function", function: { name: "list_repos", description: "List repositories for a user or the authenticated user.", parameters: { type: "object", properties: { username: { type: "string", description: "GitHub username (omit for your own repos)" }, sort: { type: "string", description: "Sort by: created, updated, pushed, full_name (default: updated)" }, perPage: { type: "integer", description: "Results per page (default 30)" } } } } },
    { type: "function", function: { name: "get_repo", description: "Get detailed information about a specific repository.", parameters: { type: "object", properties: { owner: { type: "string", description: "Repository owner" }, repo: { type: "string", description: "Repository name" } }, required: ["owner", "repo"] } } },
    { type: "function", function: { name: "create_repo", description: "Create a new GitHub repository.", parameters: { type: "object", properties: { name: { type: "string", description: "Repository name" }, description: { type: "string", description: "Repository description" }, isPrivate: { type: "boolean", description: "Make repository private (default: false)" }, autoInit: { type: "boolean", description: "Initialize with README (default: true)" } }, required: ["name"] } } },
    { type: "function", function: { name: "list_issues", description: "List issues for a repository.", parameters: { type: "object", properties: { owner: { type: "string", description: "Repository owner" }, repo: { type: "string", description: "Repository name" }, state: { type: "string", description: "Issue state: open, closed, all (default: open)" }, labels: { type: "string", description: "Comma-separated label names" }, perPage: { type: "integer", description: "Results per page (default 30)" } }, required: ["owner", "repo"] } } },
    { type: "function", function: { name: "create_issue", description: "Create a new issue in a repository.", parameters: { type: "object", properties: { owner: { type: "string", description: "Repository owner" }, repo: { type: "string", description: "Repository name" }, title: { type: "string", description: "Issue title" }, body: { type: "string", description: "Issue body (markdown)" }, labels: { type: "array", items: { type: "string" }, description: "Label names" }, assignees: { type: "array", items: { type: "string" }, description: "Assignee usernames" } }, required: ["owner", "repo", "title"] } } },
    { type: "function", function: { name: "update_issue", description: "Update an existing issue (title, body, state, labels, assignees).", parameters: { type: "object", properties: { owner: { type: "string", description: "Repository owner" }, repo: { type: "string", description: "Repository name" }, issueNumber: { type: "integer", description: "Issue number" }, title: { type: "string", description: "New title" }, body: { type: "string", description: "New body" }, state: { type: "string", description: "New state: open or closed" }, labels: { type: "array", items: { type: "string" }, description: "Labels to set" }, assignees: { type: "array", items: { type: "string" }, description: "Assignees to set" } }, required: ["owner", "repo", "issueNumber"] } } },
    { type: "function", function: { name: "list_pull_requests", description: "List pull requests for a repository.", parameters: { type: "object", properties: { owner: { type: "string", description: "Repository owner" }, repo: { type: "string", description: "Repository name" }, state: { type: "string", description: "PR state: open, closed, all (default: open)" }, perPage: { type: "integer", description: "Results per page (default 30)" } }, required: ["owner", "repo"] } } },
    { type: "function", function: { name: "get_pull_request", description: "Get full details of a specific pull request.", parameters: { type: "object", properties: { owner: { type: "string", description: "Repository owner" }, repo: { type: "string", description: "Repository name" }, pullNumber: { type: "integer", description: "Pull request number" } }, required: ["owner", "repo", "pullNumber"] } } },
    { type: "function", function: { name: "create_pull_request", description: "Create a new pull request.", parameters: { type: "object", properties: { owner: { type: "string", description: "Repository owner" }, repo: { type: "string", description: "Repository name" }, title: { type: "string", description: "PR title" }, body: { type: "string", description: "PR description" }, head: { type: "string", description: "Branch containing changes" }, base: { type: "string", description: "Branch to merge into" } }, required: ["owner", "repo", "title", "head", "base"] } } },
    { type: "function", function: { name: "merge_pull_request", description: "Merge a pull request.", parameters: { type: "object", properties: { owner: { type: "string", description: "Repository owner" }, repo: { type: "string", description: "Repository name" }, pullNumber: { type: "integer", description: "Pull request number" }, mergeMethod: { type: "string", description: "Merge method: merge, squash, rebase (default: merge)" }, commitMessage: { type: "string", description: "Custom merge commit message" } }, required: ["owner", "repo", "pullNumber"] } } },
    { type: "function", function: { name: "list_branches", description: "List branches in a repository.", parameters: { type: "object", properties: { owner: { type: "string", description: "Repository owner" }, repo: { type: "string", description: "Repository name" }, perPage: { type: "integer", description: "Results per page (default 30)" } }, required: ["owner", "repo"] } } },
    { type: "function", function: { name: "create_branch", description: "Create a new branch from an existing branch.", parameters: { type: "object", properties: { owner: { type: "string", description: "Repository owner" }, repo: { type: "string", description: "Repository name" }, branchName: { type: "string", description: "New branch name" }, fromBranch: { type: "string", description: "Source branch (default: main)" } }, required: ["owner", "repo", "branchName"] } } },
    { type: "function", function: { name: "get_file_content", description: "Get the content of a file from a repository.", parameters: { type: "object", properties: { owner: { type: "string", description: "Repository owner" }, repo: { type: "string", description: "Repository name" }, filePath: { type: "string", description: "Path to the file" }, ref: { type: "string", description: "Branch or commit SHA (default: default branch)" } }, required: ["owner", "repo", "filePath"] } } },
    { type: "function", function: { name: "create_or_update_file", description: "Create or update a file in a repository.", parameters: { type: "object", properties: { owner: { type: "string", description: "Repository owner" }, repo: { type: "string", description: "Repository name" }, filePath: { type: "string", description: "Path for the file" }, content: { type: "string", description: "File content" }, message: { type: "string", description: "Commit message" }, branch: { type: "string", description: "Target branch" }, sha: { type: "string", description: "SHA of file being replaced (required for updates)" } }, required: ["owner", "repo", "filePath", "content", "message"] } } },
    { type: "function", function: { name: "search_repos", description: "Search GitHub repositories by keyword, language, stars, etc.", parameters: { type: "object", properties: { query: { type: "string", description: "Search query (e.g. 'react language:javascript stars:>1000')" }, sort: { type: "string", description: "Sort by: stars, forks, updated (default: stars)" }, perPage: { type: "integer", description: "Results per page (default 20)" } }, required: ["query"] } } },
    { type: "function", function: { name: "search_code", description: "Search code across GitHub repositories.", parameters: { type: "object", properties: { query: { type: "string", description: "Code search query (e.g. 'useState repo:facebook/react')" }, perPage: { type: "integer", description: "Results per page (default 20)" } }, required: ["query"] } } },
    { type: "function", function: { name: "list_commits", description: "List recent commits for a repository.", parameters: { type: "object", properties: { owner: { type: "string", description: "Repository owner" }, repo: { type: "string", description: "Repository name" }, sha: { type: "string", description: "Branch name or commit SHA" }, perPage: { type: "integer", description: "Results per page (default 20)" } }, required: ["owner", "repo"] } } },
    { type: "function", function: { name: "get_user_profile", description: "Get a GitHub user's profile. Omit username to get your own profile.", parameters: { type: "object", properties: { username: { type: "string", description: "GitHub username (omit for your own)" } } } } },
    { type: "function", function: { name: "list_notifications", description: "List your GitHub notifications.", parameters: { type: "object", properties: { all: { type: "boolean", description: "Show all including read (default: false)" }, perPage: { type: "integer", description: "Results per page (default 20)" } } } } },
    { type: "function", function: { name: "list_gists", description: "List your GitHub gists.", parameters: { type: "object", properties: { perPage: { type: "integer", description: "Results per page (default 20)" } } } } }
];

// ============================================================
//  TOOL EXECUTION ROUTERS
// ============================================================

// Gmail tool names for fast lookup
const gmailToolNames = new Set(gmailTools.map(t => t.function.name));
const calendarToolNames = new Set(calendarTools.map(t => t.function.name));
const githubToolNames = new Set(githubTools.map(t => t.function.name));

async function executeGmailTool(toolName, args) {
    if (!gmailClient) throw new Error('Gmail not connected. Please authenticate first.');
    const toolMap = {
        send_email: sendEmail, search_emails: searchEmails, read_email: readEmail,
        list_emails: listEmails, trash_email: trashEmail, modify_labels: modifyLabels,
        create_draft: createDraft, reply_to_email: replyToEmail, forward_email: forwardEmail,
        list_labels: listLabels, create_label: createLabel, delete_label: deleteLabel,
        mark_as_read: markAsRead, mark_as_unread: markAsUnread, star_email: starEmail,
        unstar_email: unstarEmail, archive_email: archiveEmail, untrash_email: untrashEmail,
        get_thread: getThread, list_drafts: listDrafts, delete_draft: deleteDraft,
        send_draft: sendDraft, get_attachment_info: getAttachmentInfo, get_profile: getProfile,
        batch_modify_emails: batchModifyEmails
    };
    const fn = toolMap[toolName];
    if (!fn) throw new Error(`Unknown Gmail tool: ${toolName}`);
    return await fn(args);
}

async function executeCalendarTool(toolName, args) {
    if (!calendarClient) throw new Error('Calendar not connected. Please authenticate with Google first.');
    const toolMap = {
        list_events: listEvents, get_event: getEvent, create_event: createEvent,
        update_event: updateEvent, delete_event: deleteEvent, list_calendars: listCalendars,
        create_calendar: createCalendar, quick_add_event: quickAddEvent,
        get_free_busy: getFreeBusy, list_recurring_instances: listRecurringInstances,
        move_event: moveEvent, update_event_attendees: updateEventAttendees,
        get_calendar_colors: getCalendarColors, clear_calendar: clearCalendar,
        watch_events: watchEvents
    };
    const fn = toolMap[toolName];
    if (!fn) throw new Error(`Unknown Calendar tool: ${toolName}`);
    try {
        return await fn(args);
    } catch (error) {
        const permissionError = getCalendarPermissionError(error);
        if (permissionError) {
            calendarClient = null;
            throw new Error(permissionError);
        }
        throw error;
    }
}

async function executeGitHubTool(toolName, args) {
    if (!octokitClient) throw new Error('GitHub not connected. Please add your Personal Access Token first.');
    const toolMap = {
        list_repos: listRepos, get_repo: getRepo, create_repo: createRepo,
        list_issues: listIssues, create_issue: createIssue, update_issue: updateIssue,
        list_pull_requests: listPullRequests, get_pull_request: getPullRequest,
        create_pull_request: createPullRequest, merge_pull_request: mergePullRequest,
        list_branches: listBranches, create_branch: createBranch,
        get_file_content: getFileContent, create_or_update_file: createOrUpdateFile,
        search_repos: searchRepos, search_code: searchCode, list_commits: listCommits,
        get_user_profile: getUserProfile, list_notifications: listNotifications,
        list_gists: listGists
    };
    const fn = toolMap[toolName];
    if (!fn) throw new Error(`Unknown GitHub tool: ${toolName}`);
    return await fn(args);
}

// Master dispatcher
async function executeTool(toolName, args) {
    console.log(`[Tool] ${toolName}`, JSON.stringify(args).slice(0, 200));
    if (gmailToolNames.has(toolName)) return await executeGmailTool(toolName, args);
    if (calendarToolNames.has(toolName)) return await executeCalendarTool(toolName, args);
    if (githubToolNames.has(toolName)) return await executeGitHubTool(toolName, args);
    throw new Error(`Unknown tool: ${toolName}`);
}

// ============================================================
//  API ROUTES
// ============================================================

// Return tools list for the UI grouped by service
app.get('/api/tools', (req, res) => {
    const gmail = { service: 'gmail', connected: !!gmailClient, tools: gmailTools.map(t => ({ function: t.function })) };
    const calendarConnected = !!calendarClient && hasCalendarScope();
    const calendar = { service: 'calendar', connected: calendarConnected, tools: calendarTools.map(t => ({ function: t.function })) };
    const github = { service: 'github', connected: !!octokitClient, tools: githubTools.map(t => ({ function: t.function })) };
    res.json({ services: [gmail, calendar, github], totalTools: gmailTools.length + calendarTools.length + githubTools.length });
});

// Gmail authentication status
app.get('/api/gmail/status', (req, res) => {
    const hasCredentials = (process.env.GMAIL_CLIENT_ID || process.env.GOOGLE_CLIENT_ID) &&
        (process.env.GMAIL_CLIENT_SECRET || process.env.GOOGLE_CLIENT_SECRET);
    const hasToken = fs.existsSync(TOKEN_PATH);
    res.json({
        credentialsConfigured: !!hasCredentials,
        authenticated: hasToken && gmailClient !== null,
        toolCount: gmailTools.length
    });
});

// Calendar authentication status (same as Gmail since same OAuth)
app.get('/api/calendar/status', (req, res) => {
    const hasCredentials = (process.env.GMAIL_CLIENT_ID || process.env.GOOGLE_CLIENT_ID) &&
        (process.env.GMAIL_CLIENT_SECRET || process.env.GOOGLE_CLIENT_SECRET);
    const hasToken = fs.existsSync(TOKEN_PATH);
    const calendarScopeGranted = hasCalendarScope();
    res.json({
        credentialsConfigured: !!hasCredentials,
        authenticated: hasToken && calendarClient !== null && calendarScopeGranted,
        hasCalendarScope: calendarScopeGranted,
        requiresReconnect: hasToken && !calendarScopeGranted,
        toolCount: calendarTools.length
    });
});

// Calendar connect - triggers re-auth with calendar scope
app.get('/api/calendar/connect', (req, res) => {
    if (!oauth2Client) {
        const initialized = initOAuthClient();
        if (!initialized) {
            return res.status(400).json({ error: 'Google OAuth credentials not configured in .env file', setupRequired: true });
        }
    }
    const authUrl = oauth2Client.generateAuthUrl({ access_type: 'offline', scope: SCOPES, prompt: 'consent' });
    res.json({ authUrl });
});

// GitHub authentication status
app.get('/api/github/status', (req, res) => {
    res.json({
        authenticated: octokitClient !== null,
        toolCount: githubTools.length
    });
});

// GitHub connect with PAT
app.post('/api/github/connect', async (req, res) => {
    const { token } = req.body;
    if (!token) return res.status(400).json({ error: 'Token is required' });

    try {
        const testClient = new Octokit({ auth: token });
        const user = await testClient.rest.users.getAuthenticated();
        octokitClient = testClient;

        // Save token
        const tokenDir = path.dirname(GITHUB_TOKEN_PATH);
        if (!fs.existsSync(tokenDir)) fs.mkdirSync(tokenDir, { recursive: true });
        fs.writeFileSync(GITHUB_TOKEN_PATH, JSON.stringify({ token }, null, 2));

        res.json({ success: true, username: user.data.login, message: `Connected as ${user.data.login}` });
    } catch (error) {
        res.status(401).json({ error: `Invalid token: ${error.message}` });
    }
});

// GitHub disconnect
app.post('/api/github/disconnect', (req, res) => {
    octokitClient = null;
    if (fs.existsSync(GITHUB_TOKEN_PATH)) {
        fs.unlinkSync(GITHUB_TOKEN_PATH);
    }
    res.json({ success: true, message: 'GitHub disconnected' });
});

// Get OAuth URL
app.get('/api/gmail/auth', (req, res) => {
    if (!oauth2Client) {
        const initialized = initOAuthClient();
        if (!initialized) {
            return res.status(400).json({ error: 'Google OAuth credentials not configured in .env file', setupRequired: true });
        }
    }
    const authUrl = oauth2Client.generateAuthUrl({ access_type: 'offline', scope: SCOPES, prompt: 'consent' });
    res.json({ authUrl });
});

// OAuth callback
app.get('/oauth2callback', async (req, res) => {
    const { code } = req.query;
    if (!code) return res.status(400).send('No authorization code provided');

    try {
        const { tokens } = await oauth2Client.getToken(code);
        oauth2Client.setCredentials(tokens);
        const tokenDir = path.dirname(TOKEN_PATH);
        if (!fs.existsSync(tokenDir)) fs.mkdirSync(tokenDir, { recursive: true });
        fs.writeFileSync(TOKEN_PATH, JSON.stringify(tokens, null, 2));
        gmailClient = google.gmail({ version: 'v1', auth: oauth2Client });
        calendarClient = tokenHasScope(tokens, CALENDAR_SCOPE)
            ? google.calendar({ version: 'v3', auth: oauth2Client })
            : null;
        const calendarMessage = calendarClient
            ? 'Gmail + Calendar are ready.'
            : 'Gmail is ready. Calendar permission is still missing, so reconnect Calendar from the app.';
        res.send(`<html><body style="background:#0f0f1a;color:#fff;display:flex;align-items:center;justify-content:center;height:100vh;font-family:sans-serif"><div style="text-align:center"><h1>Google Connected!</h1><p>${calendarMessage} You can close this window.</p></div></body></html>`);
    } catch (error) {
        console.error('OAuth callback error:', error);
        res.status(500).send(`Authentication failed: ${error.message}`);
    }
});

// ============================================================
//  AGENTIC CHAT ENDPOINT — Robust multi-turn tool loop
// ============================================================
app.post('/api/chat', async (req, res) => {
    const { message, history = [] } = req.body;

    if (!process.env.OPENAI_API_KEY) {
        return res.status(500).json({ error: 'OpenAI API key missing' });
    }

    try {
        // Dynamically assemble available tools
        const availableTools = [];
        if (gmailClient) availableTools.push(...gmailTools);
        const calendarConnected = !!calendarClient && hasCalendarScope();
        if (calendarConnected) availableTools.push(...calendarTools);
        if (octokitClient) availableTools.push(...githubTools);

        const connectedServices = [];
        if (gmailClient) connectedServices.push('Gmail (25 tools)');
        if (calendarConnected) connectedServices.push('Google Calendar (15 tools)');
        if (octokitClient) connectedServices.push('GitHub (20 tools)');
        const statusText = connectedServices.length > 0 ? connectedServices.join(', ') : 'No services connected';

        const systemPrompt = `You are a powerful AI assistant with up to 60 tools across Gmail, Google Calendar, and GitHub. You can perform complex, multi-step operations across all connected services.

Connected Services: ${statusText}
Total Tools Available: ${availableTools.length}

## CORE RULES — Follow these STRICTLY:

1. **ALWAYS SEARCH FIRST**: When the user refers to an email by description, you MUST use search_emails first. When referring to repos/issues, use list/search tools first. NEVER guess IDs.

2. **CHAIN TOOLS FOR COMPLEX TASKS**: Break down complex requests into steps and execute them one by one.
   - "Read and reply to John's latest email" → search_emails → read_email → reply_to_email
   - "Create an issue for the bug and email the team" → create_issue → send_email
   - "Find calendar events for today and email the attendees" → list_events → send_email
   - "Archive all unread emails from newsletters" → search_emails → batch_modify_emails
   - "Check my PRs and create calendar events for reviews" → list_pull_requests → create_event

3. **NEVER STOP MID-TASK**: If a task requires multiple steps, keep going until FULLY completed. Do NOT ask to confirm intermediate steps unless the action is destructive.

4. **USE BATCH OPERATIONS**: When modifying multiple emails, prefer batch_modify_emails.

5. **CROSS-SERVICE OPERATIONS**: You can combine Gmail, Calendar, and GitHub tools in a single task. Think creatively about how services work together.

6. **BE PROACTIVE**: Mention relevant info you notice while processing (unread counts, upcoming events, open PRs).

## TOOL USAGE TIPS:
- Gmail: search_emails supports full Gmail query syntax (from:, to:, subject:, is:unread, has:attachment, etc.)
- Calendar: Use list_events with timeMin/timeMax for date ranges. quick_add_event for natural language.
- GitHub: Use owner/repo format. search_repos for discovery. list_issues/list_pull_requests for project management.`;

        const messages = [
            { role: 'system', content: systemPrompt },
            ...history,
            { role: 'user', content: message }
        ];

        let response = await openai.chat.completions.create({
            model: 'gpt-4o',
            messages,
            tools: availableTools.length > 0 ? availableTools : undefined,
            tool_choice: availableTools.length > 0 ? 'auto' : undefined
        });

        let assistantMessage = response.choices[0].message;
        let allToolResults = [];
        let allSteps = [];

        const MAX_TURNS = 15;
        let turnCount = 0;

        while (assistantMessage.tool_calls && turnCount < MAX_TURNS) {
            turnCount++;
            messages.push(assistantMessage);

            const toolPromises = assistantMessage.tool_calls.map(async (toolCall) => {
                const toolName = toolCall.function.name;
                let args;
                try {
                    args = JSON.parse(toolCall.function.arguments);
                } catch (e) {
                    return { toolCall, error: `Invalid arguments: ${e.message}` };
                }

                const step = { tool: toolName, args, turn: turnCount, timestamp: Date.now() };

                try {
                    const result = await executeTool(toolName, args);
                    step.result = result;
                    step.success = true;
                    return { toolCall, result, step };
                } catch (error) {
                    step.error = error.message;
                    step.success = false;
                    return { toolCall, error: error.message, step };
                }
            });

            const toolResults = await Promise.all(toolPromises);

            for (const { toolCall, result, error, step } of toolResults) {
                if (step) allSteps.push(step);

                const resultObj = error
                    ? { tool: toolCall.function.name, error }
                    : { tool: toolCall.function.name, result };
                allToolResults.push(resultObj);

                messages.push({
                    role: 'tool',
                    tool_call_id: toolCall.id,
                    content: JSON.stringify(error ? { error } : result)
                });
            }

            response = await openai.chat.completions.create({
                model: 'gpt-4o',
                messages,
                tools: availableTools.length > 0 ? availableTools : undefined,
                tool_choice: availableTools.length > 0 ? 'auto' : undefined
            });

            assistantMessage = response.choices[0].message;
        }

        let finalResponse = assistantMessage.content || '';
        if (turnCount >= MAX_TURNS && assistantMessage.tool_calls) {
            finalResponse += '\n\n(Reached maximum steps for this request. Some operations may still be pending.)';
        }

        res.json({
            response: finalResponse,
            toolResults: allToolResults,
            steps: allSteps,
            turnsUsed: turnCount
        });

    } catch (error) {
        console.error('Chat error:', error);
        res.status(500).json({ error: error.message });
    }
});

// ============================================================
//  START SERVER
// ============================================================
const credentialsConfigured = initOAuthClient();
initGitHubClient();

const totalTools = gmailTools.length + calendarTools.length + githubTools.length;
app.listen(PORT, () => {
    console.log(`\nAI Agent Server running at http://localhost:${PORT}`);
    console.log(`Total tools available: ${totalTools} (Gmail: ${gmailTools.length}, Calendar: ${calendarTools.length}, GitHub: ${githubTools.length})`);
    console.log(`Gmail: ${gmailClient ? 'Connected' : 'Not connected'}`);
    console.log(`Calendar: ${calendarClient ? 'Connected' : 'Not connected'}`);
    console.log(`GitHub: ${octokitClient ? 'Connected' : 'Not connected'}`);
});
