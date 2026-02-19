require('dotenv').config();
const express = require('express');
const cors = require('cors');
const { google } = require('googleapis');
const OpenAI = require('openai');
const { Octokit } = require('octokit');
const { Client: McpClient } = require('@modelcontextprotocol/sdk/client');
const { StdioClientTransport } = require('@modelcontextprotocol/sdk/client/stdio.js');
const { Readable } = require('stream');
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const multer = require('multer');
const WebSocket = require('ws');
const rateLimit = require('express-rate-limit');

const app = express();
const PORT = process.env.PORT || 3000;
const HOST = String(process.env.HOST || '127.0.0.1').trim() || '127.0.0.1';
const ALLOW_REMOTE_API = /^(1|true|yes)$/i.test(String(process.env.ALLOW_REMOTE_API || '').trim());

function isLoopbackAddress(address) {
    const value = String(address || '').trim().toLowerCase();
    if (!value) return false;
    if (value === '::1' || value === '127.0.0.1' || value === '::ffff:127.0.0.1') return true;
    if (value.startsWith('::ffff:127.')) return true;
    return false;
}

const CORS_TRUSTED_EXTENSION_ORIGIN_PATTERN = /^chrome-extension:\/\/[a-p]{32}$/i;

function isAllowedCorsOrigin(origin) {
    const raw = String(origin || '').trim();
    if (!raw) return true;

    const allowedOrigins = new Set([
        `http://localhost:${PORT}`,
        `http://127.0.0.1:${PORT}`,
        'http://localhost:3000',
        'http://127.0.0.1:3000'
    ]);

    if (allowedOrigins.has(raw)) return true;
    return CORS_TRUSTED_EXTENSION_ORIGIN_PATTERN.test(raw);
}

// Security: restrict CORS to localhost UI + trusted extension origins.
app.use(cors({
    origin: (origin, callback) => {
        if (isAllowedCorsOrigin(origin)) return callback(null, true);
        callback(new Error('Not allowed by CORS'));
    },
    credentials: true
}));

// Security: local-only API by default. Set ALLOW_REMOTE_API=true to disable this guard.
app.use((req, res, next) => {
    if (ALLOW_REMOTE_API) return next();
    const ip = req.ip || req.socket?.remoteAddress || '';
    if (isLoopbackAddress(ip)) return next();
    return res.status(403).json({ error: 'Remote access is disabled. Set ALLOW_REMOTE_API=true to enable.' });
});

// Security headers
app.use((req, res, next) => {
    res.setHeader('X-Content-Type-Options', 'nosniff');
    res.setHeader('X-Frame-Options', 'DENY');
    res.setHeader('X-XSS-Protection', '1; mode=block');
    res.setHeader('Referrer-Policy', 'strict-origin-when-cross-origin');
    res.setHeader('Permissions-Policy', 'camera=(), microphone=(self), geolocation=()');
    next();
});

// Rate limiting
const apiLimiter = rateLimit({ windowMs: 60 * 1000, max: 30, message: { error: 'Too many requests, please try again later.' } });
const chatLimiter = rateLimit({ windowMs: 60 * 1000, max: 15, message: { error: 'Too many requests, please try again later.' } });
const uploadLimiter = rateLimit({ windowMs: 60 * 1000, max: 10, message: { error: 'Too many uploads, please try again later.' } });
// Apply rate limiting (skip status endpoints — they're read-only polling)
app.use('/api/', (req, res, next) => {
    if (req.path.endsWith('/status')) return next();
    return apiLimiter(req, res, next);
});
app.use('/api/chat', chatLimiter);
app.use('/api/upload', uploadLimiter);

const JSON_BODY_LIMIT = String(process.env.JSON_BODY_LIMIT || '5gb').trim() || '5gb';
app.use(express.json({ limit: JSON_BODY_LIMIT }));
app.use(express.urlencoded({ extended: false, limit: JSON_BODY_LIMIT }));
app.use(express.static('public'));

// File upload configuration
const UPLOADS_DIR = path.join(__dirname, 'uploads');

// Security: validate file paths are within UPLOADS_DIR to prevent path traversal
function validateLocalPath(filePath) {
    if (!filePath) return false;
    const resolved = path.resolve(filePath);
    const uploadsResolved = path.resolve(UPLOADS_DIR);
    return resolved.startsWith(uploadsResolved + path.sep) || resolved === uploadsResolved;
}
if (!fs.existsSync(UPLOADS_DIR)) {
    fs.mkdirSync(UPLOADS_DIR, { recursive: true });
}

const storage = multer.diskStorage({
    destination: (req, file, cb) => {
        cb(null, UPLOADS_DIR);
    },
    filename: (req, file, cb) => {
        const uniqueSuffix = `${Date.now()}-${crypto.randomBytes(6).toString('hex')}`;
        const ext = path.extname(file.originalname);
        const basename = path.basename(file.originalname, ext);
        cb(null, `${basename}-${uniqueSuffix}${ext}`);
    }
});

const MAX_UPLOAD_BYTES = Math.max(
    1 * 1024 * 1024,
    Number.parseInt(process.env.MAX_UPLOAD_BYTES || `${5 * 1024 * 1024 * 1024}`, 10) || (5 * 1024 * 1024 * 1024)
);

const upload = multer({
    storage,
    limits: {
        fileSize: MAX_UPLOAD_BYTES
    },
    fileFilter: (req, file, cb) => {
        // Allow all file types for now
        cb(null, true);
    }
});

// Store uploaded files metadata temporarily (in-memory)
const uploadedFiles = new Map(); // fileId -> { path, originalName, size, mimeType, uploadedAt }

// Clean up old uploads every 10 minutes
// Covers both user-uploaded files (tracked in uploadedFiles Map) AND
// Drive/GCS files downloaded as email attachment intermediates (not in Map)
setInterval(() => {
    const MAX_AGE = 15 * 60 * 1000; // 15 minutes
    const now = Date.now();

    // 1. Clean tracked user uploads
    for (const [fileId, fileInfo] of uploadedFiles.entries()) {
        if (now - fileInfo.uploadedAt > MAX_AGE) {
            try {
                if (fs.existsSync(fileInfo.path)) fs.unlinkSync(fileInfo.path);
                uploadedFiles.delete(fileId);
            } catch (error) {
                console.error('Failed to clean up tracked upload:', error);
            }
        }
    }

    // 2. Clean any untracked files in uploads/ dir (Drive/GCS downloads for attachments)
    try {
        const files = fs.readdirSync(UPLOADS_DIR);
        for (const filename of files) {
            const filepath = path.join(UPLOADS_DIR, filename);
            try {
                const stat = fs.statSync(filepath);
                if (stat.isFile() && (now - stat.mtimeMs) > MAX_AGE) {
                    fs.unlinkSync(filepath);
                }
            } catch (_) { }
        }
    } catch (error) {
        console.error('Failed to sweep uploads dir:', error);
    }
}, 10 * 60 * 1000); // Run every 10 minutes

// OpenAI setup
const openai = new OpenAI({
    apiKey: process.env.OPENAI_API_KEY
});
const DEFAULT_OPENAI_MODEL = 'gpt-4o-mini';
const OPENAI_MODEL = String(process.env.OPENAI_MODEL || DEFAULT_OPENAI_MODEL).trim() || DEFAULT_OPENAI_MODEL;
const OPENAI_FALLBACK_MODEL = String(process.env.OPENAI_FALLBACK_MODEL || '').trim();
const OPENAI_CHAT_MAX_RETRIES = Math.min(
    4,
    Math.max(0, Number.parseInt(process.env.OPENAI_CHAT_MAX_RETRIES || '2', 10) || 2)
);
const parsedOpenAiTemperature = Number.parseFloat(process.env.OPENAI_TEMPERATURE ?? '0.2');
const OPENAI_TEMPERATURE = Number.isFinite(parsedOpenAiTemperature)
    ? Math.min(1, Math.max(0, parsedOpenAiTemperature))
    : 0.2;
const parsedOpenAiMaxOutput = Number.parseInt(process.env.OPENAI_MAX_OUTPUT_TOKENS || '', 10);
const OPENAI_MAX_OUTPUT_TOKENS = Number.isFinite(parsedOpenAiMaxOutput) && parsedOpenAiMaxOutput > 0
    ? parsedOpenAiMaxOutput
    : 1500;
const CHAT_HISTORY_MAX_MESSAGES = Math.max(
    6,
    Number.parseInt(process.env.CHAT_HISTORY_MAX_MESSAGES || '14', 10) || 14
);
const CHAT_HISTORY_MAX_MESSAGE_CHARS = Math.max(
    600,
    Number.parseInt(process.env.CHAT_HISTORY_MAX_MESSAGE_CHARS || '1800', 10) || 1800
);
const CHAT_HISTORY_MAX_TOTAL_CHARS = Math.max(
    3000,
    Number.parseInt(process.env.CHAT_HISTORY_MAX_TOTAL_CHARS || '12000', 10) || 12000
);
const MODEL_TOOL_RESULT_MAX_CHARS = Math.max(
    1200,
    Number.parseInt(process.env.MODEL_TOOL_RESULT_MAX_CHARS || '2500', 10) || 2500
);
const MODEL_TOOL_VALUE_MAX_STRING_CHARS = Math.max(
    400,
    Number.parseInt(process.env.MODEL_TOOL_VALUE_MAX_STRING_CHARS || '900', 10) || 900
);
const MODEL_TOOL_VALUE_MAX_ARRAY_ITEMS = Math.max(
    8,
    Number.parseInt(process.env.MODEL_TOOL_VALUE_MAX_ARRAY_ITEMS || '20', 10) || 20
);
const MODEL_TOOL_VALUE_MAX_OBJECT_KEYS = Math.max(
    20,
    Number.parseInt(process.env.MODEL_TOOL_VALUE_MAX_OBJECT_KEYS || '50', 10) || 50
);
const ASSISTANT_RESPONSE_MAX_CHARS = 12000;
const DRIVE_TEXT_EDIT_MAX_CHARS = 250000;
const MAX_CHAT_MESSAGE_CHARS = Math.max(
    200,
    Number.parseInt(process.env.MAX_CHAT_MESSAGE_CHARS || '12000', 10) || 12000
);
const MAX_CHAT_HISTORY_ITEMS = Math.max(
    2,
    Number.parseInt(process.env.MAX_CHAT_HISTORY_ITEMS || '40', 10) || 40
);
const MAX_CHAT_ATTACHED_FILES = Math.max(
    1,
    Number.parseInt(process.env.MAX_CHAT_ATTACHED_FILES || '10', 10) || 10
);

// Google OAuth scopes setup (Gmail, Calendar, Chat, Drive, Sheets, Docs)
const SCOPES = [
    'https://www.googleapis.com/auth/gmail.readonly',
    'https://www.googleapis.com/auth/gmail.send',
    'https://www.googleapis.com/auth/gmail.modify',
    'https://www.googleapis.com/auth/gmail.labels',
    'https://www.googleapis.com/auth/gmail.compose',
    'https://www.googleapis.com/auth/gmail.settings.basic',
    'https://www.googleapis.com/auth/calendar',
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/documents',
    'https://www.googleapis.com/auth/chat.spaces.readonly',
    'https://www.googleapis.com/auth/chat.messages.create',
    'https://www.googleapis.com/auth/chat.messages.readonly'
];
const CALENDAR_SCOPE = 'https://www.googleapis.com/auth/calendar';
const DRIVE_SCOPE = 'https://www.googleapis.com/auth/drive';
const SHEETS_SCOPE = 'https://www.googleapis.com/auth/spreadsheets';
const DOCS_SCOPE = 'https://www.googleapis.com/auth/documents';
const GCHAT_REQUIRED_SCOPES = [
    'https://www.googleapis.com/auth/chat.spaces.readonly',
    'https://www.googleapis.com/auth/chat.messages.create',
    'https://www.googleapis.com/auth/chat.messages.readonly'
];
const GOOGLE_OAUTH_PROMPT = 'consent select_account';
const GITHUB_OAUTH_SCOPES = ['repo', 'read:user', 'user:email', 'notifications', 'gist'];
const GITHUB_OAUTH_STATE_TTL_MS = 10 * 60 * 1000;
const GOOGLE_OAUTH_STATE_TTL_MS = 10 * 60 * 1000;
const SHEETS_MCP_TOOL_PREFIX = 'sheets_mcp__';
const SHEETS_MCP_ENABLED = String(process.env.SHEETS_MCP_ENABLED || 'true').toLowerCase() !== 'false';
const SHEETS_MCP_COMMAND = process.env.SHEETS_MCP_COMMAND || 'node';
const SHEETS_MCP_ARGS = process.env.SHEETS_MCP_ARGS
    ? process.env.SHEETS_MCP_ARGS.split(',').map(arg => arg.trim()).filter(Boolean)
    : [path.join(__dirname, 'node_modules', '@isaacphi', 'mcp-gdrive', 'dist', 'index.js')];
const SHEETS_MCP_CREDS_DIR = process.env.SHEETS_MCP_CREDS_DIR ||
    path.join(process.env.USERPROFILE || process.env.HOME || '.', '.gmail-mcp', 'sheets-mcp');

// Microsoft Outlook OAuth2
const OUTLOOK_SCOPES = [
    'openid',
    'profile',
    'offline_access',
    'User.Read',
    'Mail.Read',
    'Mail.ReadWrite',
    'Mail.Send',
    'Team.ReadBasic.All',
    'Channel.ReadBasic.All',
    'ChannelMessage.Read.All',
    'ChannelMessage.Send',
    'Chat.Read',
    'Chat.ReadWrite',
    'ChatMessage.Send'
];
const TEAMS_REQUIRED_SCOPES = ['Team.ReadBasic.All', 'Channel.ReadBasic.All', 'Chat.Read'];
const OUTLOOK_OAUTH_STATE_TTL_MS = 10 * 60 * 1000;
const OUTLOOK_AUTHORITY = 'https://login.microsoftonline.com/common';
const OUTLOOK_GRAPH_BASE = 'https://graph.microsoft.com/v1.0';

const TOKEN_PATH = path.join(process.env.USERPROFILE || process.env.HOME, '.gmail-mcp', 'token.json');
const GITHUB_TOKEN_PATH = path.join(process.env.USERPROFILE || process.env.HOME, '.gmail-mcp', 'github-token.json');
const OUTLOOK_TOKEN_PATH = path.join(process.env.USERPROFILE || process.env.HOME, '.gmail-mcp', 'outlook-token.json');
const SCHEDULED_TASKS_PATH = path.join(process.env.USERPROFILE || process.env.HOME, '.gmail-mcp', 'scheduled-tasks.json');

let oauth2Client = null;
let gmailClient = null;
let calendarClient = null;
let gchatClient = null;
let driveClient = null;
let sheetsClient = null;
let docsClient = null;
let octokitClient = null;
let githubUsername = null;
let githubAuthMethod = null;
let sheetsMcpClient = null;
let sheetsMcpTransport = null;
let sheetsMcpTools = [];
let sheetsMcpError = null;
const githubOAuthStateStore = new Map();
const googleOAuthStateStore = new Map();
let cachedPrimaryEmail = null;
let scheduledTasks = [];
let schedulerInterval = null;
let schedulerTickInProgress = false;

// Outlook state
let outlookAccessToken = null;
let outlookRefreshToken = null;
let outlookTokenExpiry = null;
let outlookUserEmail = null;
let outlookAuthMethod = 'oauth';
const outlookOAuthStateStore = new Map();

// GCS (Cloud Storage) state
let gcsClient = null;
let gcsAuthenticated = false;
let gcsProjectId = null;

const MEETING_TRANSCRIPTION_TOOL_MAX_RESULTS = Math.max(
    5,
    Number.parseInt(process.env.MEETING_TRANSCRIPTION_TOOL_MAX_RESULTS || '50', 10) || 50
);
const MEETING_TRANSCRIPTION_TEXT_MAX_CHARS = Math.max(
    50000,
    Number.parseInt(process.env.MEETING_TRANSCRIPTION_TEXT_MAX_CHARS || '500000', 10) || 500000
);
const MEETING_TRANSCRIPTION_CHUNK_CHARS = Math.max(
    12000,
    Number.parseInt(process.env.MEETING_TRANSCRIPTION_CHUNK_CHARS || '45000', 10) || 45000
);
const MEETING_TRANSCRIPTION_SINGLE_PASS_CHARS = Math.max(
    30000,
    Number.parseInt(process.env.MEETING_TRANSCRIPTION_SINGLE_PASS_CHARS || '120000', 10) || 120000
);

const EMAIL_REGEX = /[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/ig;
const DISALLOWED_ATTENDEE_DOMAINS = new Set(['example.com', 'test.com', 'domain.com', 'email.com']);
const NO_REPLY_PATTERN = /(no-?reply|do-?not-?reply|noreply)/i;

function parseScopes(scopeString) {
    if (!scopeString || typeof scopeString !== 'string') return new Set();
    return new Set(scopeString.split(/\s+/).filter(Boolean));
}

function tokenHasScope(token, scope) {
    if (!token || !scope) return false;
    return parseScopes(token.scope).has(scope);
}

function tokenHasScopes(token, scopes) {
    if (!token || !Array.isArray(scopes) || scopes.length === 0) return false;
    const granted = parseScopes(token.scope);
    return scopes.every(scope => granted.has(scope));
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

function hasGchatScopes() {
    if (oauth2Client && tokenHasScopes(oauth2Client.credentials, GCHAT_REQUIRED_SCOPES)) {
        return true;
    }
    const savedToken = readSavedToken();
    return tokenHasScopes(savedToken, GCHAT_REQUIRED_SCOPES);
}

function hasDriveScope() {
    if (oauth2Client && tokenHasScope(oauth2Client.credentials, DRIVE_SCOPE)) {
        return true;
    }
    const savedToken = readSavedToken();
    return tokenHasScope(savedToken, DRIVE_SCOPE);
}

function hasSheetsScope() {
    if (oauth2Client && tokenHasScope(oauth2Client.credentials, SHEETS_SCOPE)) {
        return true;
    }
    const savedToken = readSavedToken();
    return tokenHasScope(savedToken, SHEETS_SCOPE);
}

function hasDocsScope() {
    if (oauth2Client && tokenHasScope(oauth2Client.credentials, DOCS_SCOPE)) {
        return true;
    }
    const savedToken = readSavedToken();
    return tokenHasScope(savedToken, DOCS_SCOPE);
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

function getGchatPermissionError(error) {
    const status = error?.code || error?.status || error?.response?.status;
    const message = String(error?.message || '');
    const looksLikeScopeError = /insufficient|permission|scope|forbidden|unauthorized/i.test(message);
    if ((status === 401 || status === 403) && looksLikeScopeError) {
        return 'Google Chat permission is missing or expired. Please reconnect Google Chat from the Chat panel and try again.';
    }
    return null;
}

function getDrivePermissionError(error) {
    const status = error?.code || error?.status || error?.response?.status;
    const message = String(error?.message || '');
    const looksLikeScopeError = /insufficient|permission|scope|forbidden|unauthorized/i.test(message);
    if ((status === 401 || status === 403) && looksLikeScopeError) {
        return 'Google Drive permission is missing or expired. Please reconnect Google Drive from the Drive panel and try again.';
    }
    return null;
}

function getSheetsPermissionError(error) {
    const status = error?.code || error?.status || error?.response?.status;
    const message = String(error?.message || '');
    const looksLikeScopeError = /insufficient|permission|scope|forbidden|unauthorized/i.test(message);
    if ((status === 401 || status === 403) && looksLikeScopeError) {
        return 'Google Sheets permission is missing or expired. Please reconnect Google Sheets from the Sheets panel and try again.';
    }
    return null;
}

function formatLocalDate(date) {
    const y = date.getFullYear();
    const m = String(date.getMonth() + 1).padStart(2, '0');
    const d = String(date.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
}

function getCurrentDateContext() {
    const now = new Date();
    const tomorrow = new Date(now);
    tomorrow.setDate(tomorrow.getDate() + 1);
    const yesterday = new Date(now);
    yesterday.setDate(yesterday.getDate() - 1);

    const timeZone = Intl.DateTimeFormat().resolvedOptions().timeZone || 'UTC';
    const weekday = now.toLocaleDateString('en-US', { weekday: 'long' });

    return {
        nowIso: now.toISOString(),
        today: formatLocalDate(now),
        tomorrow: formatLocalDate(tomorrow),
        yesterday: formatLocalDate(yesterday),
        weekday,
        timeZone
    };
}

function isGitHubOAuthConfigured() {
    const clientId = process.env.GITHUB_CLIENT_ID;
    const clientSecret = process.env.GITHUB_CLIENT_SECRET;
    if (!clientId || !clientSecret) return false;
    if (clientId.includes('your_github_client_id_here')) return false;
    if (clientSecret.includes('your_github_client_secret_here')) return false;
    return true;
}

function getGitHubRedirectUri() {
    return process.env.GITHUB_REDIRECT_URI || `http://localhost:${PORT}/github/callback`;
}

function ensureParentDirectory(filePath) {
    const dir = path.dirname(filePath);
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
}

function makeId(prefix = 'id') {
    if (typeof crypto.randomUUID === 'function') {
        return `${prefix}_${crypto.randomUUID()}`;
    }
    return `${prefix}_${Date.now()}_${Math.random().toString(36).slice(2, 10)}`;
}

function isValidDailyTime(value) {
    return /^\d{2}:\d{2}$/.test(String(value || '')) &&
        Number(value.slice(0, 2)) >= 0 &&
        Number(value.slice(0, 2)) <= 23 &&
        Number(value.slice(3, 5)) >= 0 &&
        Number(value.slice(3, 5)) <= 59;
}

function dailyTimeToMinutes(value) {
    if (!isValidDailyTime(value)) return null;
    return (Number(value.slice(0, 2)) * 60) + Number(value.slice(3, 5));
}

function clampPagination(value, { fallback = 20, min = 1, max = 100 } = {}) {
    const parsed = Number.parseInt(value, 10);
    if (!Number.isFinite(parsed)) return fallback;
    return Math.min(max, Math.max(min, parsed));
}

function truncateText(value, maxChars = 2000) {
    const text = String(value ?? '');
    if (text.length <= maxChars) return text;
    if (maxChars <= 20) return text.slice(0, Math.max(0, maxChars));
    return `${text.slice(0, maxChars - 20)}\n...[truncated]`;
}

const EMAIL_SEND_CONFIRMATION_TOOLS = new Set([
    'send_email',
    'reply_to_email',
    'forward_email',
    'send_draft',
    'outlook_send_email',
    'outlook_reply_to_email',
    'outlook_forward_email',
    'outlook_send_draft'
]);

const SEND_TO_RECIPIENT_TOOLS = new Set([
    'send_email',
    'forward_email',
    'outlook_send_email',
    'outlook_forward_email'
]);

const SEND_CC_BCC_RECIPIENT_TOOLS = new Set([
    'send_email',
    'outlook_send_email'
]);

function isTruthyConfirmationValue(value) {
    if (value === true) return true;
    if (typeof value === 'number') return value === 1;
    if (typeof value !== 'string') return false;
    const normalized = value.trim().toLowerCase();
    return normalized === 'true' || normalized === 'yes' || normalized === 'y' || normalized === 'confirm' || normalized === 'confirmed' || normalized === 'ok';
}

function hasEmailSendConfirmation(args) {
    if (!args || typeof args !== 'object') return false;
    const confirmationCandidates = [
        args.confirmSend,
        args.confirm_send,
        args.userConfirmed,
        args.confirmed,
        args.approved
    ];
    return confirmationCandidates.some(isTruthyConfirmationValue);
}

function isAffirmativeEmailSendReply(value) {
    const text = String(value || '').trim().toLowerCase();
    if (!text) return false;
    if (/\b(cancel|stop|dont|don't|not now|wait)\b/.test(text)) return false;
    return /\b(yes|yep|yeah|confirm|confirmed|looks good|all good|proceed|go ahead|send it|send now|send|ok)\b/.test(text);
}

function historyContainsEmailSendConfirmationPrompt(history = []) {
    if (!Array.isArray(history)) return false;
    return history.some(item => {
        if (!item || item.role !== 'assistant' || typeof item.content !== 'string') return false;
        return item.content.toLowerCase().includes('email send confirmation required before sending');
    });
}

function isEmailSendConfirmedForTurn({ message, history = [] }) {
    return historyContainsEmailSendConfirmationPrompt(history) && isAffirmativeEmailSendReply(message);
}

function sanitizeEmailPreviewText(value, maxChars = 240) {
    const plain = String(value ?? '')
        .replace(/<[^>]+>/g, ' ')
        .replace(/\s+/g, ' ')
        .trim();
    return truncateText(plain, maxChars);
}

function normalizeRecipientPreview(value) {
    if (!value) return '';
    if (Array.isArray(value)) {
        return value.map(item => String(item || '').trim()).filter(Boolean).join(', ');
    }
    return String(value).trim();
}

function buildEmailSendConfirmationMessage(toolName, args = {}) {
    const to = normalizeRecipientPreview(args.to);
    const cc = normalizeRecipientPreview(args.cc);
    const bcc = normalizeRecipientPreview(args.bcc);
    const subject = sanitizeEmailPreviewText(args.subject || '(no subject)', 180);
    const bodySource = args.body ?? args.comment ?? args.additionalMessage ?? '';
    const bodyPreview = sanitizeEmailPreviewText(bodySource || '(empty body)', 320);
    const draftId = args.draftId ? String(args.draftId) : '';
    const messageId = args.messageId ? String(args.messageId) : '';

    const lines = [
        'Email send confirmation required before sending.',
        '',
        'Please review:',
        to ? `- To (resolved email): ${to}` : '- To (resolved email): (not provided)',
        cc ? `- CC (resolved email): ${cc}` : '',
        bcc ? `- BCC (resolved email): ${bcc}` : '',
        `- Subject: ${subject}`
    ].filter(Boolean);

    if (toolName.includes('reply')) {
        lines.push(messageId ? `- Replying to message ID: ${messageId}` : '- Replying to: (message ID missing)');
    } else if (toolName.includes('forward')) {
        lines.push(messageId ? `- Forwarding message ID: ${messageId}` : '- Forwarding: (message ID missing)');
    } else if (toolName.includes('draft')) {
        if (draftId) {
            lines.push(`- Draft ID: ${draftId}`);
        } else if (messageId) {
            lines.push(`- Draft message ID: ${messageId}`);
        } else {
            lines.push('- Draft ID: (not provided)');
        }
    }

    lines.push(`- Body preview: ${bodyPreview}`);

    return lines.join('\n');
}

function stripSendConfirmationFlags(args) {
    if (!args || typeof args !== 'object') return args;
    const nextArgs = { ...args };
    delete nextArgs.confirmSend;
    delete nextArgs.confirm_send;
    delete nextArgs.userConfirmed;
    delete nextArgs.confirmed;
    delete nextArgs.approved;
    return nextArgs;
}

async function prepareEmailSendArgs(toolName, args = {}) {
    const normalizedArgs = (args && typeof args === 'object') ? { ...args } : {};

    if (SEND_TO_RECIPIENT_TOOLS.has(toolName)) {
        normalizedArgs.to = await normalizeMessageRecipients(normalizedArgs.to, {
            fieldName: 'to',
            requireAtLeastOne: true
        });
    }

    if (SEND_CC_BCC_RECIPIENT_TOOLS.has(toolName)) {
        normalizedArgs.cc = await normalizeMessageRecipients(normalizedArgs.cc, { fieldName: 'cc' });
        normalizedArgs.bcc = await normalizeMessageRecipients(normalizedArgs.bcc, { fieldName: 'bcc' });
    }

    return normalizedArgs;
}

function shouldRequestUserInputForToolError(errorMessage) {
    const text = String(errorMessage || '').toLowerCase();
    if (!text) return false;
    if (text.includes('email send confirmation required')) return true;
    if (text.includes('please provide exact email address')) return true;
    if (text.includes('could not resolve') && text.includes('email')) return true;
    if (text.includes('no valid') && text.includes('email')) return true;
    return false;
}

function sanitizeHistoryText(value) {
    return String(value ?? '')
        .replace(/\u0000/g, '')
        .trim();
}

function buildSafeChatHistory(
    history = [],
    {
        maxMessages = CHAT_HISTORY_MAX_MESSAGES,
        maxMessageChars = CHAT_HISTORY_MAX_MESSAGE_CHARS,
        maxTotalChars = CHAT_HISTORY_MAX_TOTAL_CHARS
    } = {}
) {
    const normalized = (Array.isArray(history) ? history : [])
        .filter(item =>
            item &&
            (item.role === 'user' || item.role === 'assistant') &&
            typeof item.content === 'string'
        )
        .map(item => ({
            role: item.role,
            content: truncateText(sanitizeHistoryText(item.content), maxMessageChars)
        }))
        .filter(item => item.content.length > 0);

    const selected = [];
    let totalChars = 0;
    for (let i = normalized.length - 1; i >= 0; i -= 1) {
        if (selected.length >= maxMessages) break;
        const item = normalized[i];
        const remaining = maxTotalChars - totalChars;
        if (remaining <= 80) break;

        let content = item.content;
        if (content.length > remaining) {
            content = truncateText(content, remaining);
        }
        selected.unshift({ role: item.role, content });
        totalChars += content.length;
    }

    return selected;
}

function cloneToolsForRequest(tools = []) {
    if (!Array.isArray(tools)) return [];
    return tools.map(tool => {
        try {
            return JSON.parse(JSON.stringify(tool));
        } catch {
            const params = tool?.function?.parameters;
            return {
                ...tool,
                function: tool?.function
                    ? {
                        ...tool.function,
                        parameters: params && typeof params === 'object'
                            ? {
                                ...params,
                                properties: params.properties && typeof params.properties === 'object'
                                    ? { ...params.properties }
                                    : params.properties
                            }
                            : params
                    }
                    : tool?.function
            };
        }
    });
}

function applyEmailSignatureHintToTools(tools = [], userFirstName = '') {
    if (!Array.isArray(tools) || !userFirstName) return;
    const sigNote = ` Always end email body with "Best regards,\\n${userFirstName}". No other signature or placeholders.`;
    const emailSendTools = new Set([
        'send_email',
        'reply_to_email',
        'forward_email',
        'create_draft',
        'outlook_send_email',
        'outlook_reply_to_email',
        'outlook_forward_email',
        'outlook_create_draft'
    ]);

    for (const tool of tools) {
        if (!emailSendTools.has(tool?.function?.name)) continue;
        const bodySchema = tool?.function?.parameters?.properties?.body;
        if (!bodySchema || typeof bodySchema.description !== 'string') continue;
        if (bodySchema.description.includes('Best regards,')) continue;
        bodySchema.description += sigNote;
    }
}

function normalizeReadableText(text, maxChars = 40000) {
    let normalized = String(text ?? '')
        .replace(/\r\n?/g, '\n')
        .replace(/\u0000/g, '');

    normalized = normalized
        .replace(/[^\x09\x0A\x0D\x20-\x7E\u00A0-\u024F\u0370-\u03FF\u0400-\u04FF]/g, ' ')
        .replace(/[ \t]+\n/g, '\n')
        .replace(/[ \t]{2,}/g, ' ')
        .replace(/\n{3,}/g, '\n\n')
        .trim();

    const originalLength = normalized.length;
    const truncated = originalLength > maxChars;
    if (truncated) {
        normalized = `${normalized.slice(0, Math.max(0, maxChars - 20))}\n...[truncated]`;
    }

    return { text: normalized, truncated, originalLength };
}

function safeLocalDateOnly(value) {
    if (!value) return '';
    const parsed = new Date(value);
    if (Number.isNaN(parsed.getTime())) return '';
    return formatLocalDate(parsed);
}

function parseDateRangeInput(dateValue) {
    const raw = String(dateValue || '').trim();
    if (!raw) return null;
    if (!/^\d{4}-\d{2}-\d{2}$/.test(raw)) {
        throw new Error('date must be in YYYY-MM-DD format');
    }
    const start = new Date(`${raw}T00:00:00`);
    if (Number.isNaN(start.getTime())) {
        throw new Error('Invalid date value');
    }
    const end = new Date(start);
    end.setDate(end.getDate() + 1);
    return {
        date: raw,
        startIso: start.toISOString(),
        endIso: end.toISOString()
    };
}

function splitTextIntoSizedChunks(text, maxChars = 45000) {
    const source = String(text || '').trim();
    if (!source) return [];
    const size = Math.max(2000, Number.parseInt(maxChars, 10) || 45000);
    const chunks = [];
    let cursor = 0;

    while (cursor < source.length) {
        const remaining = source.length - cursor;
        if (remaining <= size) {
            chunks.push(source.slice(cursor).trim());
            break;
        }

        let splitAt = source.lastIndexOf('\n\n', cursor + size);
        if (splitAt <= cursor + (size * 0.5)) {
            splitAt = source.lastIndexOf('\n', cursor + size);
        }
        if (splitAt <= cursor + (size * 0.5)) {
            splitAt = source.lastIndexOf(' ', cursor + size);
        }
        if (splitAt <= cursor) {
            splitAt = cursor + size;
        }

        const chunk = source.slice(cursor, splitAt).trim();
        if (chunk) chunks.push(chunk);
        cursor = splitAt;
    }

    return chunks.filter(Boolean);
}

function normalizeMeetingSummaryPayload(payload) {
    const source = (payload && typeof payload === 'object' && !Array.isArray(payload)) ? payload : {};

    const normalizeStringList = (value, fallback = []) => {
        if (Array.isArray(value)) {
            return value
                .map(item => String(item || '').trim())
                .filter(Boolean)
                .slice(0, 20);
        }
        const single = String(value || '').trim();
        if (single) return [single];
        return fallback;
    };

    const actionItemsRaw = Array.isArray(source.actionItems) ? source.actionItems : [];
    const actionItems = actionItemsRaw
        .map(item => {
            if (!item) return null;
            if (typeof item === 'string') {
                return {
                    owner: 'TBD',
                    action: item.trim(),
                    due: 'TBD'
                };
            }
            if (typeof item === 'object' && !Array.isArray(item)) {
                const owner = String(item.owner || item.assignee || item.person || 'TBD').trim() || 'TBD';
                const action = String(item.action || item.task || '').trim();
                const due = String(item.due || item.dueDate || item.deadline || 'TBD').trim() || 'TBD';
                if (!action) return null;
                return { owner, action, due };
            }
            return null;
        })
        .filter(Boolean)
        .slice(0, 30);

    return {
        summary: normalizeStringList(source.summary, ['No clear discussion summary was identified.']),
        actionItems,
        nextSteps: normalizeStringList(source.nextSteps, ['No next steps identified.'])
    };
}

function renderMeetingSummaryMarkdown({ summary = [], actionItems = [], nextSteps = [], sourceLink }) {
    const summaryLines = summary.length > 0
        ? summary.map(item => `- ${item}`)
        : ['- No clear discussion summary was identified.'];
    const actionLines = actionItems.length > 0
        ? actionItems.map(item => `- ${item.owner || 'TBD'} | ${item.action || 'TBD'} | ${item.due || 'TBD'}`)
        : ['- None identified.'];
    const nextStepLines = nextSteps.length > 0
        ? nextSteps.map(item => `- ${item}`)
        : ['- None identified.'];

    const sections = [
        '## Summary',
        ...summaryLines,
        '',
        '## Action Items (Owner | Action | Due)',
        ...actionLines,
        '',
        '## Next Steps',
        ...nextStepLines
    ];

    if (sourceLink) {
        sections.push('', `Source File: ${sourceLink}`);
    }

    return sections.join('\n');
}

async function summarizeMeetingTranscriptChunkWithOpenAI({
    title,
    transcriptText,
    chunkLabel = ''
}) {
    const prompt = `Meeting title: ${title || 'Meeting'}
${chunkLabel ? `Section: ${chunkLabel}` : ''}

Transcript:
${transcriptText}

Return ONLY valid JSON with this shape:
{
  "summary": ["point 1", "point 2"],
  "actionItems": [{"owner":"Name or TBD","action":"Task","due":"Date or TBD"}],
  "nextSteps": ["step 1", "step 2"]
}

Rules:
- Be concise and factual.
- Keep each list item short.
- If a field is unknown, use "TBD".
- Do not include markdown or commentary.`;

    const completion = await openai.chat.completions.create({
        model: OPENAI_MODEL,
        messages: [
            {
                role: 'system',
                content: 'You extract structured meeting notes from transcripts. Return JSON only.'
            },
            {
                role: 'user',
                content: prompt
            }
        ],
        temperature: 0.15,
        max_tokens: 1400
    });

    const raw = completion?.choices?.[0]?.message?.content || '';
    const parsed = parseJsonObjectFromText(raw);
    if (!parsed) {
        throw new Error('Failed to parse meeting summary JSON from OpenAI response.');
    }

    return normalizeMeetingSummaryPayload(parsed);
}

function mergeMeetingSummaryPayloads(parts = []) {
    const summaryMap = new Map();
    const nextStepMap = new Map();
    const mergedActions = [];

    const addTextToMap = (target, text) => {
        const value = String(text || '').trim();
        if (!value) return;
        const key = value.toLowerCase();
        if (target.has(key)) return;
        target.set(key, value);
    };

    for (const part of parts) {
        const normalized = normalizeMeetingSummaryPayload(part);
        normalized.summary.forEach(item => addTextToMap(summaryMap, item));
        normalized.nextSteps.forEach(item => addTextToMap(nextStepMap, item));
        normalized.actionItems.forEach(item => {
            const owner = String(item.owner || 'TBD').trim() || 'TBD';
            const action = String(item.action || '').trim();
            const due = String(item.due || 'TBD').trim() || 'TBD';
            if (!action) return;
            const dedupeKey = `${owner.toLowerCase()}|${action.toLowerCase()}`;
            if (mergedActions.some(existing => `${existing.owner.toLowerCase()}|${existing.action.toLowerCase()}` === dedupeKey)) {
                return;
            }
            mergedActions.push({ owner, action, due });
        });
    }

    return {
        summary: Array.from(summaryMap.values()).slice(0, 20),
        actionItems: mergedActions.slice(0, 30),
        nextSteps: Array.from(nextStepMap.values()).slice(0, 20)
    };
}

async function refineMergedMeetingSummaryWithOpenAI({ title, mergedPayload }) {
    const normalized = normalizeMeetingSummaryPayload(mergedPayload);
    const prompt = `Meeting title: ${title || 'Meeting'}

Input summary data (JSON):
${JSON.stringify(normalized)}

Return ONLY valid JSON with this same shape:
{
  "summary": ["point 1", "point 2"],
  "actionItems": [{"owner":"Name or TBD","action":"Task","due":"Date or TBD"}],
  "nextSteps": ["step 1", "step 2"]
}

Rules:
- Remove duplicates.
- Keep concise, factual points.
- Preserve all important action items.
- Use "TBD" for unknown owner/due values.`;

    const completion = await openai.chat.completions.create({
        model: OPENAI_MODEL,
        messages: [
            {
                role: 'system',
                content: 'You clean and consolidate structured meeting notes. Return JSON only.'
            },
            {
                role: 'user',
                content: prompt
            }
        ],
        temperature: 0.1,
        max_tokens: 1400
    });

    const raw = completion?.choices?.[0]?.message?.content || '';
    const parsed = parseJsonObjectFromText(raw);
    if (!parsed) {
        return normalized;
    }
    return normalizeMeetingSummaryPayload(parsed);
}

function isLikelyTextMimeType(mimeType) {
    const mime = String(mimeType || '').toLowerCase();
    if (!mime) return false;
    if (mime.startsWith('text/')) return true;
    return [
        'json',
        'xml',
        'yaml',
        'csv',
        'javascript',
        'x-sh',
        'markdown',
        'sql'
    ].some(token => mime.includes(token));
}

function hasHighBinaryRatio(buffer, sampleSize = 4096) {
    if (!buffer || buffer.length === 0) return false;
    const sample = buffer.subarray(0, Math.min(sampleSize, buffer.length));
    let binaryCount = 0;
    for (const byte of sample) {
        const isPrintable = (byte >= 32 && byte <= 126) || byte === 9 || byte === 10 || byte === 13;
        if (!isPrintable) binaryCount += 1;
    }
    return (binaryCount / sample.length) > 0.3;
}

function extractPrintableStringsFromBuffer(buffer, { minLength = 4, maxChars = 40000 } = {}) {
    if (!buffer || buffer.length === 0) return '';
    const sampled = buffer.subarray(0, Math.min(buffer.length, maxChars * 4));
    const latinText = sampled.toString('latin1');
    const regex = new RegExp(`[\\x20-\\x7E]{${Math.max(2, Number(minLength) || 4)},}`, 'g');
    const matches = latinText.match(regex) || [];
    if (matches.length === 0) return '';
    const joined = matches.slice(0, 500).join('\n');
    return normalizeReadableText(joined, maxChars).text;
}

function schedulerDateParts(now = new Date()) {
    const year = now.getFullYear();
    const month = String(now.getMonth() + 1).padStart(2, '0');
    const day = String(now.getDate()).padStart(2, '0');
    const hour = String(now.getHours()).padStart(2, '0');
    const minute = String(now.getMinutes()).padStart(2, '0');
    return {
        date: `${year}-${month}-${day}`,
        hhmm: `${hour}:${minute}`,
        timestamp: now.toISOString()
    };
}

function sanitizeScheduledTask(task, { forCreate = false } = {}) {
    const normalized = {
        id: String(task.id || '').trim() || makeId('task'),
        name: String(task.name || '').trim() || 'Untitled task',
        instruction: String(task.instruction || '').trim(),
        time: String(task.time || '').trim(),
        enabled: task.enabled !== undefined ? !!task.enabled : true,
        createdAt: task.createdAt || new Date().toISOString(),
        updatedAt: new Date().toISOString(),
        lastRunAt: task.lastRunAt || null,
        lastRunDate: task.lastRunDate || null,
        lastStatus: task.lastStatus || 'never',
        lastError: task.lastError || '',
        lastResponse: task.lastResponse || ''
    };

    if (!normalized.instruction) {
        throw new Error('instruction is required');
    }
    if (!isValidDailyTime(normalized.time)) {
        throw new Error('time must be in HH:MM 24-hour format');
    }

    if (!forCreate) {
        normalized.createdAt = task.createdAt || normalized.createdAt;
    }
    return normalized;
}

function loadScheduledTasksFromDisk() {
    try {
        if (!fs.existsSync(SCHEDULED_TASKS_PATH)) {
            scheduledTasks = [];
            return;
        }
        const raw = fs.readFileSync(SCHEDULED_TASKS_PATH, 'utf8');
        const parsed = JSON.parse(raw);
        const list = Array.isArray(parsed) ? parsed : [];
        scheduledTasks = list
            .map(task => {
                try {
                    return sanitizeScheduledTask(task);
                } catch (error) {
                    console.warn(`Skipping invalid scheduled task: ${error.message}`);
                    return null;
                }
            })
            .filter(Boolean);
    } catch (error) {
        console.error('Failed to load scheduled tasks:', error.message);
        scheduledTasks = [];
    }
}

function saveScheduledTasksToDisk() {
    ensureParentDirectory(SCHEDULED_TASKS_PATH);
    fs.writeFileSync(SCHEDULED_TASKS_PATH, JSON.stringify(scheduledTasks, null, 2));
}

function saveGitHubTokenData(data) {
    const tokenDir = path.dirname(GITHUB_TOKEN_PATH);
    if (!fs.existsSync(tokenDir)) fs.mkdirSync(tokenDir, { recursive: true });
    fs.writeFileSync(GITHUB_TOKEN_PATH, JSON.stringify(data, null, 2));
}

function clearGitHubAuth() {
    octokitClient = null;
    githubUsername = null;
    githubAuthMethod = null;
    if (fs.existsSync(GITHUB_TOKEN_PATH)) fs.unlinkSync(GITHUB_TOKEN_PATH);
}

// Outlook OAuth helpers
function isOutlookOAuthConfigured() {
    const clientId = process.env.OUTLOOK_CLIENT_ID;
    const clientSecret = process.env.OUTLOOK_CLIENT_SECRET;
    if (!clientId || !clientSecret) return false;
    if (clientId.includes('your_outlook_client_id_here')) return false;
    if (clientSecret.includes('your_outlook_client_secret_here')) return false;
    return true;
}

function getOutlookRedirectUri() {
    return process.env.OUTLOOK_REDIRECT_URI || `http://localhost:${PORT}/outlook/callback`;
}

function saveOutlookTokenData(data) {
    const tokenDir = path.dirname(OUTLOOK_TOKEN_PATH);
    if (!fs.existsSync(tokenDir)) fs.mkdirSync(tokenDir, { recursive: true });
    fs.writeFileSync(OUTLOOK_TOKEN_PATH, JSON.stringify(data, null, 2));
}

function clearOutlookAuth() {
    outlookAccessToken = null;
    outlookRefreshToken = null;
    outlookTokenExpiry = null;
    outlookUserEmail = null;
    if (fs.existsSync(OUTLOOK_TOKEN_PATH)) fs.unlinkSync(OUTLOOK_TOKEN_PATH);
}

function hasTeamsScopes() {
    try {
        if (!fs.existsSync(OUTLOOK_TOKEN_PATH)) return false;
        const data = JSON.parse(fs.readFileSync(OUTLOOK_TOKEN_PATH, 'utf8'));
        if (!data.scope) return false;
        const granted = new Set(data.scope.split(/\s+/).filter(Boolean));
        return TEAMS_REQUIRED_SCOPES.every(s => granted.has(s));
    } catch {
        return false;
    }
}

function getGoogleOAuthCredentials() {
    const clientId = process.env.GMAIL_CLIENT_ID || process.env.GOOGLE_CLIENT_ID || '';
    const clientSecret = process.env.GMAIL_CLIENT_SECRET || process.env.GOOGLE_CLIENT_SECRET || '';
    return { clientId, clientSecret };
}

function isGoogleOAuthConfigured() {
    const { clientId, clientSecret } = getGoogleOAuthCredentials();
    if (!clientId || !clientSecret) return false;
    const idLower = clientId.toLowerCase();
    const secretLower = clientSecret.toLowerCase();
    if (idLower.includes('your_google_client_id') || idLower.includes('your_gmail_client_id')) return false;
    if (secretLower.includes('your_google_client_secret') || secretLower.includes('your_gmail_client_secret')) return false;
    return true;
}

function ensureSheetsMcpOAuthKeyfile(clientId, clientSecret) {
    if (!clientId || !clientSecret) {
        throw new Error('Google OAuth credentials are required to initialize Sheets MCP');
    }

    fs.mkdirSync(SHEETS_MCP_CREDS_DIR, { recursive: true });
    const keyfilePath = path.join(SHEETS_MCP_CREDS_DIR, 'gcp-oauth.keys.json');
    if (fs.existsSync(keyfilePath)) return keyfilePath;

    const desktopOAuthConfig = {
        installed: {
            client_id: clientId,
            project_id: process.env.GOOGLE_PROJECT_ID || 'gmail-mcp-chat',
            auth_uri: 'https://accounts.google.com/o/oauth2/auth',
            token_uri: 'https://oauth2.googleapis.com/token',
            auth_provider_x509_cert_url: 'https://www.googleapis.com/oauth2/v1/certs',
            client_secret: clientSecret,
            redirect_uris: ['http://localhost']
        }
    };
    fs.writeFileSync(keyfilePath, JSON.stringify(desktopOAuthConfig, null, 2));
    return keyfilePath;
}

function toOpenAiToolFromMcpTool(tool) {
    return {
        type: 'function',
        function: {
            name: `${SHEETS_MCP_TOOL_PREFIX}${tool.name}`,
            description: `[Sheets MCP] ${tool.description || `Run ${tool.name}`}`,
            parameters: tool.inputSchema || { type: 'object', properties: {} }
        }
    };
}

function getSheetsMcpConnectionError() {
    if (!sheetsMcpError) {
        return 'Google Sheets MCP is not connected.';
    }
    return `Google Sheets MCP is not connected: ${sheetsMcpError}`;
}

function parseMcpTextContent(content) {
    if (!Array.isArray(content)) return [];
    return content
        .filter(part => part && part.type === 'text' && typeof part.text === 'string')
        .map(part => part.text);
}

async function initSheetsMcpClient() {
    if (!SHEETS_MCP_ENABLED) {
        sheetsMcpError = 'Disabled by SHEETS_MCP_ENABLED=false';
        return;
    }
    if (sheetsMcpClient) return;

    const { clientId, clientSecret } = getGoogleOAuthCredentials();
    if (!clientId || !clientSecret) {
        sheetsMcpError = 'Missing GOOGLE_CLIENT_ID/GOOGLE_CLIENT_SECRET for Sheets MCP.';
        return;
    }

    try {
        ensureSheetsMcpOAuthKeyfile(clientId, clientSecret);

        const client = new McpClient({
            name: 'gmail-mcp-chat-sheets-bridge',
            version: '1.0.0'
        });
        const transport = new StdioClientTransport({
            command: SHEETS_MCP_COMMAND,
            args: SHEETS_MCP_ARGS,
            env: {
                ...process.env,
                CLIENT_ID: clientId,
                CLIENT_SECRET: clientSecret,
                GDRIVE_CREDS_DIR: SHEETS_MCP_CREDS_DIR
            },
            stderr: 'pipe'
        });

        if (transport.stderr) {
            transport.stderr.on('data', chunk => {
                const text = String(chunk || '').trim();
                if (text) console.log(`[Sheets MCP] ${text}`);
            });
        }

        await client.connect(transport);
        const listResult = await client.listTools();
        const discoveredTools = Array.isArray(listResult?.tools) ? listResult.tools : [];

        sheetsMcpClient = client;
        sheetsMcpTransport = transport;
        sheetsMcpTools = discoveredTools.map(toOpenAiToolFromMcpTool);
        sheetsMcpError = null;
        console.log(`Sheets MCP connected with ${sheetsMcpTools.length} tool(s).`);
    } catch (error) {
        sheetsMcpClient = null;
        sheetsMcpTransport = null;
        sheetsMcpTools = [];
        sheetsMcpError = error?.message || 'Unknown Sheets MCP connection failure';
        console.error('Sheets MCP init error:', sheetsMcpError);
    }
}

async function executeSheetsMcpTool(toolName, args) {
    if (!sheetsMcpClient) {
        throw new Error(getSheetsMcpConnectionError());
    }
    if (!toolName.startsWith(SHEETS_MCP_TOOL_PREFIX)) {
        throw new Error(`Unknown Sheets MCP tool: ${toolName}`);
    }
    const mcpToolName = toolName.slice(SHEETS_MCP_TOOL_PREFIX.length);
    const result = await sheetsMcpClient.callTool({
        name: mcpToolName,
        arguments: args || {}
    });
    const textBlocks = parseMcpTextContent(result?.content);
    return {
        provider: '@isaacphi/mcp-gdrive',
        tool: mcpToolName,
        isError: !!result?.isError,
        text: textBlocks.join('\n\n'),
        structuredContent: result?.structuredContent || null,
        content: result?.content || []
    };
}

async function closeSheetsMcpClient() {
    try {
        if (sheetsMcpTransport) {
            await sheetsMcpTransport.close();
        }
    } catch (error) {
        console.error('Error closing Sheets MCP transport:', error.message);
    } finally {
        sheetsMcpClient = null;
        sheetsMcpTransport = null;
        sheetsMcpTools = [];
    }
}

function pruneGoogleOAuthStates() {
    const now = Date.now();
    for (const [state, ts] of googleOAuthStateStore.entries()) {
        if (now - ts > GOOGLE_OAUTH_STATE_TTL_MS) googleOAuthStateStore.delete(state);
    }
}

function issueGoogleOAuthState() {
    pruneGoogleOAuthStates();
    const state = crypto.randomBytes(24).toString('hex');
    googleOAuthStateStore.set(state, Date.now());
    return state;
}

function consumeGoogleOAuthState(state) {
    if (!state || !googleOAuthStateStore.has(state)) return false;
    const ts = googleOAuthStateStore.get(state);
    googleOAuthStateStore.delete(state);
    return (Date.now() - ts) <= GOOGLE_OAUTH_STATE_TTL_MS;
}

function pruneGithubOAuthStates() {
    const now = Date.now();
    for (const [state, ts] of githubOAuthStateStore.entries()) {
        if (now - ts > GITHUB_OAUTH_STATE_TTL_MS) githubOAuthStateStore.delete(state);
    }
}

function issueGithubOAuthState() {
    pruneGithubOAuthStates();
    const state = crypto.randomBytes(24).toString('hex');
    githubOAuthStateStore.set(state, Date.now());
    return state;
}

function consumeGithubOAuthState(state) {
    if (!state || !githubOAuthStateStore.has(state)) return false;
    const ts = githubOAuthStateStore.get(state);
    githubOAuthStateStore.delete(state);
    return (Date.now() - ts) <= GITHUB_OAUTH_STATE_TTL_MS;
}

function pruneOutlookOAuthStates() {
    const now = Date.now();
    for (const [s, ts] of outlookOAuthStateStore) {
        if (now - ts > OUTLOOK_OAUTH_STATE_TTL_MS) outlookOAuthStateStore.delete(s);
    }
}

function issueOutlookOAuthState() {
    pruneOutlookOAuthStates();
    const state = crypto.randomBytes(24).toString('hex');
    outlookOAuthStateStore.set(state, Date.now());
    return state;
}

function consumeOutlookOAuthState(state) {
    if (!state || !outlookOAuthStateStore.has(state)) return false;
    const ts = outlookOAuthStateStore.get(state);
    outlookOAuthStateStore.delete(state);
    return (Date.now() - ts) <= OUTLOOK_OAUTH_STATE_TTL_MS;
}

// Initialize OAuth client from environment variables
function initOAuthClient() {
    const { clientId, clientSecret } = getGoogleOAuthCredentials();

    if (!isGoogleOAuthConfigured()) {
        console.log('  Google OAuth credentials not configured in .env');
        return false;
    }

    try {
        oauth2Client = new google.auth.OAuth2(
            clientId,
            clientSecret,
            `http://localhost:${PORT}/oauth2callback`
        );

        // Auto-save refreshed tokens to disk so they survive server restarts
        oauth2Client.on('tokens', (tokens) => {
            try {
                const existing = fs.existsSync(TOKEN_PATH)
                    ? JSON.parse(fs.readFileSync(TOKEN_PATH, 'utf8'))
                    : {};
                // Merge: keep existing refresh_token if new one isn't provided
                const updated = { ...existing, ...tokens };
                if (!updated.refresh_token && existing.refresh_token) {
                    updated.refresh_token = existing.refresh_token;
                }
                fs.writeFileSync(TOKEN_PATH, JSON.stringify(updated, null, 2));
                console.log('Google token refreshed and saved to disk');
            } catch (err) {
                console.error('Failed to save refreshed token:', err.message);
            }
        });

        if (fs.existsSync(TOKEN_PATH)) {
            const token = JSON.parse(fs.readFileSync(TOKEN_PATH, 'utf8'));
            oauth2Client.setCredentials(token);

            // Force refresh if token is expired (so clients work immediately on startup)
            if (token.refresh_token && token.expiry_date && Date.now() > token.expiry_date) {
                console.log('Google token expired, refreshing...');
                oauth2Client.refreshAccessToken().then(({ credentials }) => {
                    oauth2Client.setCredentials(credentials);
                    // on('tokens') listener will save to disk
                    console.log('Google token refreshed successfully on startup');
                }).catch(err => {
                    console.error('Failed to refresh Google token on startup:', err.message);
                });
            }

            gmailClient = google.gmail({ version: 'v1', auth: oauth2Client });
            calendarClient = tokenHasScope(token, CALENDAR_SCOPE)
                ? google.calendar({ version: 'v3', auth: oauth2Client })
                : null;
            gchatClient = tokenHasScopes(token, GCHAT_REQUIRED_SCOPES)
                ? google.chat({ version: 'v1', auth: oauth2Client })
                : null;
            driveClient = tokenHasScope(token, DRIVE_SCOPE)
                ? google.drive({ version: 'v3', auth: oauth2Client })
                : null;
            sheetsClient = tokenHasScope(token, SHEETS_SCOPE)
                ? google.sheets({ version: 'v4', auth: oauth2Client })
                : null;
            docsClient = tokenHasScope(token, DOCS_SCOPE)
                ? google.docs({ version: 'v1', auth: oauth2Client })
                : null;
            if (calendarClient) {
                console.log('Gmail + Calendar clients initialized with existing token');
            } else {
                console.log('Gmail initialized. Calendar scope missing in token; reconnect required for Calendar tools.');
            }
            if (!gchatClient) {
                console.log('Google Chat scopes missing in token; reconnect required for Chat tools.');
            }
            if (!driveClient) {
                console.log('Google Drive scope missing in token; reconnect required for Drive tools.');
            }
            if (!sheetsClient) {
                console.log('Google Sheets scope missing in token; reconnect required for Sheets tools.');
            }
            if (!docsClient) {
                console.log('Google Docs scope missing in token; reconnect required for Docs tools.');
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
                githubUsername = data.username || null;
                githubAuthMethod = data.authMethod || 'pat';
                console.log('GitHub client initialized with saved token');
                return true;
            }
        }
    } catch (error) {
        console.error('Error initializing GitHub client:', error);
    }
    return false;
}

// Initialize Outlook client from saved token
async function refreshOutlookToken() {
    if (!outlookRefreshToken || !isOutlookOAuthConfigured()) return false;
    try {
        const params = new URLSearchParams({
            client_id: process.env.OUTLOOK_CLIENT_ID,
            client_secret: process.env.OUTLOOK_CLIENT_SECRET,
            refresh_token: outlookRefreshToken,
            grant_type: 'refresh_token',
            scope: OUTLOOK_SCOPES.join(' ')
        });
        const resp = await fetch(`${OUTLOOK_AUTHORITY}/oauth2/v2.0/token`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
            body: params.toString()
        });
        if (!resp.ok) {
            console.error('Outlook token refresh failed:', resp.status);
            return false;
        }
        const tokenData = await resp.json();
        outlookAccessToken = tokenData.access_token;
        if (tokenData.refresh_token) outlookRefreshToken = tokenData.refresh_token;
        outlookTokenExpiry = Date.now() + (tokenData.expires_in * 1000) - 60000;
        const existingData = fs.existsSync(OUTLOOK_TOKEN_PATH) ? JSON.parse(fs.readFileSync(OUTLOOK_TOKEN_PATH, 'utf8')) : {};
        saveOutlookTokenData({
            ...existingData,
            access_token: outlookAccessToken,
            refresh_token: outlookRefreshToken,
            expiry: outlookTokenExpiry,
            email: outlookUserEmail,
            scope: tokenData.scope || existingData.scope || OUTLOOK_SCOPES.join(' '),
            connectedAt: new Date().toISOString()
        });
        return true;
    } catch (error) {
        console.error('Outlook token refresh error:', error.message);
        return false;
    }
}

async function initOutlookClient() {
    try {
        if (!fs.existsSync(OUTLOOK_TOKEN_PATH)) return false;
        const data = JSON.parse(fs.readFileSync(OUTLOOK_TOKEN_PATH, 'utf8'));
        if (!data.access_token || !data.refresh_token) return false;
        outlookAccessToken = data.access_token;
        outlookRefreshToken = data.refresh_token;
        outlookTokenExpiry = data.expiry || 0;
        outlookUserEmail = data.email || null;
        outlookAuthMethod = 'oauth';
        if (Date.now() >= outlookTokenExpiry) {
            const refreshed = await refreshOutlookToken();
            if (!refreshed) {
                clearOutlookAuth();
                return false;
            }
        }
        console.log(`Outlook client initialized (${outlookUserEmail || 'unknown user'})`);
        return true;
    } catch (error) {
        console.error('Error initializing Outlook client:', error);
        return false;
    }
}

function initGcsClient() {
    try {
        const keyFile = process.env.GCP_SERVICE_ACCOUNT_KEY_FILE;
        const projectId = process.env.GCP_PROJECT_ID;
        if (!keyFile || !projectId) {
            console.log('GCS: Skipped (GCP_SERVICE_ACCOUNT_KEY_FILE or GCP_PROJECT_ID not set)');
            return false;
        }
        if (!fs.existsSync(keyFile)) {
            console.error(`GCS: Service account key file not found: ${keyFile}`);
            return false;
        }
        const auth = new google.auth.GoogleAuth({
            keyFile,
            scopes: ['https://www.googleapis.com/auth/cloud-platform']
        });
        gcsClient = google.storage({ version: 'v1', auth });
        gcsAuthenticated = true;
        gcsProjectId = projectId;
        console.log(`GCS client initialized (project: ${projectId})`);
        return true;
    } catch (error) {
        console.error('Error initializing GCS client:', error.message);
        gcsAuthenticated = false;
        return false;
    }
}

// ============================================================
//  HELPER UTILITIES
// ============================================================

function sanitizeHeaderValue(value) {
    return String(value ?? '')
        .replace(/[\r\n]+/g, ' ')
        .trim();
}

function normalizeAddressHeader(value) {
    if (!value) return '';
    if (Array.isArray(value)) {
        return value
            .map(item => sanitizeHeaderValue(item))
            .filter(Boolean)
            .join(', ');
    }
    return sanitizeHeaderValue(value);
}

function buildRawMessage({ to, subject, body, cc, bcc, inReplyTo, references, threadId, attachments }) {
    // Format body with proper line breaks for better readability
    const formattedBody = body ? body.replace(/\n/g, '<br>') : '';

    const lines = [];
    const toHeader = normalizeAddressHeader(to);
    const ccHeader = normalizeAddressHeader(cc);
    const bccHeader = normalizeAddressHeader(bcc);
    if (toHeader) lines.push(`To: ${toHeader}`);
    if (ccHeader) lines.push(`Cc: ${ccHeader}`);
    if (bccHeader) lines.push(`Bcc: ${bccHeader}`);
    lines.push(`Subject: ${sanitizeHeaderValue(subject || '')}`);
    if (inReplyTo) lines.push(`In-Reply-To: ${sanitizeHeaderValue(inReplyTo)}`);
    if (references) lines.push(`References: ${sanitizeHeaderValue(references)}`);

    // If attachments, use multipart/mixed
    if (attachments && attachments.length > 0) {
        const boundary = `----=_Part_${Date.now()}_${crypto.randomBytes(8).toString('hex')}`;
        lines.push(`Content-Type: multipart/mixed; boundary="${boundary}"`);
        lines.push('');

        // Body part
        lines.push(`--${boundary}`);
        lines.push('Content-Type: text/html; charset=utf-8');
        lines.push('');
        lines.push(formattedBody);
        lines.push('');

        // Attachment parts
        for (const attachment of attachments) {
            lines.push(`--${boundary}`);
            lines.push(`Content-Type: ${attachment.mimeType || 'application/octet-stream'}`);
            lines.push('Content-Transfer-Encoding: base64');
            lines.push(`Content-Disposition: attachment; filename="${attachment.filename}"`);
            lines.push('');
            lines.push(attachment.content); // Already base64 encoded
            lines.push('');
        }

        lines.push(`--${boundary}--`);
    } else {
        // Simple email without attachments
        lines.push('Content-Type: text/html; charset=utf-8');
        lines.push('');
        lines.push(formattedBody);
    }

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
async function sendEmail({ to, subject, body, cc, bcc, attachments }) {
    if (!gmailClient) throw new Error('Gmail not authenticated');
    const toRecipients = await normalizeMessageRecipients(to, { fieldName: 'to', requireAtLeastOne: true });
    const ccRecipients = await normalizeMessageRecipients(cc, { fieldName: 'cc' });
    const bccRecipients = await normalizeMessageRecipients(bcc, { fieldName: 'bcc' });

    // Process attachments if provided
    let processedAttachments = [];
    if (attachments && Array.isArray(attachments)) {
        for (const att of attachments) {
            if (att.localPath) {
                if (!validateLocalPath(att.localPath)) {
                    throw new Error('Attachment path is not allowed. Only files in the uploads directory can be attached.');
                }
                if (!fs.existsSync(att.localPath)) {
                    throw new Error('Attachment file not found. Use the localPath from download_drive_file_to_local.');
                }
                // Read file from disk and encode as base64
                const fileContent = fs.readFileSync(att.localPath);
                const base64Content = fileContent.toString('base64');

                // Get MIME type from extension if not provided
                let mimeType = att.mimeType;
                if (!mimeType) {
                    const ext = path.extname(att.localPath).toLowerCase();
                    const mimeTypes = {
                        '.pdf': 'application/pdf',
                        '.doc': 'application/msword',
                        '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                        '.xls': 'application/vnd.ms-excel',
                        '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                        '.jpg': 'image/jpeg',
                        '.jpeg': 'image/jpeg',
                        '.png': 'image/png',
                        '.txt': 'text/plain',
                        '.csv': 'text/csv',
                        '.zip': 'application/zip'
                    };
                    mimeType = mimeTypes[ext] || 'application/octet-stream';
                }

                processedAttachments.push({
                    filename: att.filename || path.basename(att.localPath),
                    mimeType,
                    content: base64Content
                });
            }
        }
    }

    const raw = buildRawMessage({
        to: toRecipients,
        subject,
        body,
        cc: ccRecipients,
        bcc: bccRecipients,
        attachments: processedAttachments.length > 0 ? processedAttachments : undefined
    });

    const result = await gmailClient.users.messages.send({
        userId: 'me',
        requestBody: { raw }
    });

    const attachmentMsg = processedAttachments.length > 0
        ? ` with ${processedAttachments.length} attachment(s)`
        : '';

    return {
        success: true,
        messageId: result.data.id,
        message: `Email sent to ${toRecipients.join(', ')}${attachmentMsg}`,
        attachmentCount: processedAttachments.length
    };
}

// 2. Search Emails
async function searchEmails({ query, maxResults = 20 }) {
    if (!gmailClient) throw new Error('Gmail not authenticated');
    const limit = clampPagination(maxResults, { fallback: 20, max: 100 });
    const response = await gmailClient.users.messages.list({ userId: 'me', q: query, maxResults: limit });
    if (!response.data.messages) return { results: [], totalEstimate: 0, message: 'No emails found' };

    const emails = await Promise.all(
        response.data.messages.slice(0, limit).map(async (msg) => {
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
    const limit = clampPagination(maxResults, { fallback: 20, max: 100 });
    const response = await gmailClient.users.messages.list({ userId: 'me', labelIds: [label], maxResults: limit });
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
    const toRecipients = await normalizeMessageRecipients(to, { fieldName: 'to', requireAtLeastOne: true });
    const ccRecipients = await normalizeMessageRecipients(cc, { fieldName: 'cc' });
    const bccRecipients = await normalizeMessageRecipients(bcc, { fieldName: 'bcc' });
    const raw = buildRawMessage({ to: toRecipients, subject, body, cc: ccRecipients, bcc: bccRecipients });
    const result = await gmailClient.users.drafts.create({
        userId: 'me',
        requestBody: { message: { raw } }
    });
    return { success: true, draftId: result.data.id, message: `Draft created for ${toRecipients.join(', ')}` };
}

// 8. Reply to Email
async function replyToEmail({ messageId, body, cc }) {
    if (!gmailClient) throw new Error('Gmail not authenticated');
    const original = await gmailClient.users.messages.get({ userId: 'me', id: messageId, format: 'full' });
    const headers = original.data.payload.headers;
    const h = extractHeaders(headers, ['Subject', 'From', 'To', 'Message-ID']);

    const replyTo = (h.from || '').trim();
    if (!replyTo) {
        throw new Error('Cannot reply because the original message has no From header.');
    }
    const originalSubject = (h.subject || '').trim() || '(no subject)';
    const subject = /^re:/i.test(originalSubject) ? originalSubject : `Re: ${originalSubject}`;
    const ccRecipients = await normalizeMessageRecipients(cc, { fieldName: 'cc' });
    const raw = buildRawMessage({
        to: replyTo, subject, body, cc: ccRecipients,
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
    const originalSubject = (h.subject || '').trim() || '(no subject)';
    const subject = /^fwd?:/i.test(originalSubject) ? originalSubject : `Fwd: ${originalSubject}`;

    const toRecipients = await normalizeMessageRecipients(to, { fieldName: 'to', requireAtLeastOne: true });
    const raw = buildRawMessage({ to: toRecipients, subject, body: forwardBody });
    const result = await gmailClient.users.messages.send({ userId: 'me', requestBody: { raw } });
    return { success: true, messageId: result.data.id, message: `Email forwarded to ${toRecipients.join(', ')}` };
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
    const limit = clampPagination(maxResults, { fallback: 20, max: 100 });
    const response = await gmailClient.users.drafts.list({ userId: 'me', maxResults: limit });
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
    cachedPrimaryEmail = (profile.data.emailAddress || '').toLowerCase() || cachedPrimaryEmail;
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
//  GOOGLE CALENDAR TOOL IMPLEMENTATIONS
// ============================================================

function getMeetLinkFromEvent(eventData) {
    if (!eventData) return null;
    if (eventData.hangoutLink) return eventData.hangoutLink;
    const entryPoints = eventData.conferenceData?.entryPoints || [];
    const videoEntry = entryPoints.find(e => e.entryPointType === 'video');
    return videoEntry?.uri || null;
}

function extractEmailsFromText(text) {
    if (!text || typeof text !== 'string') return [];
    return [...new Set((text.match(EMAIL_REGEX) || []).map(e => e.toLowerCase()))];
}

function normalizeIdentity(text) {
    if (!text || typeof text !== 'string') return '';
    return text
        .toLowerCase()
        .replace(/['"]/g, '')
        .replace(/[_\-.,]+/g, ' ')
        .replace(/\s+/g, ' ')
        .trim();
}

function normalizeAttendeeDirective(text) {
    return String(text || '')
        .toLowerCase()
        .replace(/[^a-z0-9\s]/g, ' ')
        .replace(/\s+/g, ' ')
        .trim();
}

function isSelfReferenceAttendeeToken(text) {
    const normalized = normalizeAttendeeDirective(text);
    return normalized === 'me' || normalized === 'myself' || normalized === 'i';
}

function isSelfOnlyAttendeeDirective(text) {
    const normalized = normalizeAttendeeDirective(text);
    if (!normalized) return false;

    if (normalized === 'just me' || normalized === 'only me') return true;
    if (normalized === 'just myself' || normalized === 'only myself') return true;
    if (normalized === 'me only' || normalized === 'myself only') return true;
    if (normalized === 'for me only' || normalized === 'just for me') return true;
    if (normalized === 'no attendees') return true;
    return /\b(just|only)\s+(me|myself)\b/.test(normalized);
}

function hasSelfOnlyAttendeeIntentInMessage(message) {
    const normalized = normalizeAttendeeDirective(message);
    if (!normalized) return false;
    if (/\b(just|only)\s+(me|myself)\b/.test(normalized)) return true;
    if (/\b(me|myself)\s+only\b/.test(normalized)) return true;
    if (/\bjust for me\b/.test(normalized)) return true;
    return false;
}

function hasMeetingIntentInMessage(message) {
    const text = String(message || '').toLowerCase();
    if (!text) return false;
    return /\b(create|schedule|book|set up|setup)\b[\s\S]*\b(meeting|meet|call|sync|standup)\b/.test(text)
        || /\b(meeting|google meet|video call)\b/.test(text);
}

function isValidEmail(value) {
    if (!value || typeof value !== 'string') return false;
    return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(value.trim());
}

function getEmailDomain(email) {
    const i = email.lastIndexOf('@');
    return i === -1 ? '' : email.slice(i + 1).toLowerCase();
}

function isDisallowedAttendeeEmail(email) {
    if (!isValidEmail(email)) return true;
    const domain = getEmailDomain(email);
    if (DISALLOWED_ATTENDEE_DOMAINS.has(domain)) return true;
    if (NO_REPLY_PATTERN.test(email)) return true;
    return false;
}

function splitRecipientTokens(input) {
    const values = Array.isArray(input) ? input : [input];
    const tokens = [];
    for (const value of values) {
        const text = String(value ?? '').trim();
        if (!text) continue;
        if (/[<>]/.test(text)) {
            tokens.push(text);
            continue;
        }
        const parts = text.split(/[,;\n]+/).map(part => part.trim()).filter(Boolean);
        if (parts.length > 0) tokens.push(...parts);
        else tokens.push(text);
    }
    return tokens;
}

async function normalizeMessageRecipients(recipientsInput, { fieldName = 'recipient', requireAtLeastOne = false } = {}) {
    const tokens = splitRecipientTokens(recipientsInput);
    if (tokens.length === 0) {
        if (requireAtLeastOne) {
            throw new Error(`No valid ${fieldName} email addresses provided. Please provide exact email address(es).`);
        }
        return [];
    }

    const resolved = [];
    const unresolved = [];

    for (const token of tokens) {
        const raw = String(token || '').trim();
        if (!raw) continue;

        if (isValidEmail(raw) && !isDisallowedAttendeeEmail(raw)) {
            resolved.push(raw.toLowerCase());
            continue;
        }

        const extracted = extractEmailsFromText(raw).filter(email => !isDisallowedAttendeeEmail(email));
        if (extracted.length > 0) {
            resolved.push(...extracted.map(email => email.toLowerCase()));
            continue;
        }

        const found = await resolveEmailFromGmailHistory(raw);
        if (found) {
            resolved.push(found.toLowerCase());
            continue;
        }

        unresolved.push(raw);
    }

    const deduped = [...new Set(resolved)];
    if (requireAtLeastOne && deduped.length === 0 && unresolved.length === 0) {
        throw new Error(`No valid ${fieldName} email addresses could be resolved. ASK THE USER: "Could you please provide the exact email address for the recipient?".`);
    }
    if (unresolved.length > 0) {
        throw new Error(`Could not resolve email address for: ${unresolved.join(', ')}. ASK THE USER: "I couldn't find an email address for ${unresolved.join(', ')} in your Gmail history. Could you provide their exact email address?"`);
    }
    return deduped;
}

async function getPrimaryEmailAddress() {
    if (cachedPrimaryEmail) return cachedPrimaryEmail;
    if (!gmailClient) return null;
    try {
        const profile = await gmailClient.users.getProfile({ userId: 'me' });
        cachedPrimaryEmail = (profile?.data?.emailAddress || '').toLowerCase() || null;
        return cachedPrimaryEmail;
    } catch {
        return null;
    }
}

function attendeeLookupKey(input) {
    const raw = String(input || '').trim();
    if (!raw) return '';
    if (isValidEmail(raw)) {
        if (!isDisallowedAttendeeEmail(raw)) return raw.toLowerCase();
        return normalizeIdentity(raw.split('@')[0]);
    }
    return normalizeIdentity(raw);
}

async function resolveEmailFromGmailHistory(identity) {
    const key = attendeeLookupKey(identity);
    if (!key || !gmailClient) return null;

    if (isValidEmail(key) && !isDisallowedAttendeeEmail(key)) {
        return key.toLowerCase();
    }

    const ownEmail = await getPrimaryEmailAddress();
    const queries = [
        `"${key}"`,
        `from:${key}`,
        `to:${key}`
    ];
    const keyTokens = key.split(' ').filter(Boolean);
    for (const token of keyTokens) {
        if (token.length < 2) continue;
        queries.push(`"${token}"`, `from:${token}`, `to:${token}`);
    }
    const requiresStrictHeaderMatch = key.includes(' ');

    const scoreByEmail = new Map();
    const seenMessageIds = new Set();

    for (const q of queries) {
        let messages = [];
        try {
            const list = await gmailClient.users.messages.list({ userId: 'me', q, maxResults: 10 });
            messages = list?.data?.messages || [];
        } catch {
            continue;
        }

        for (const msg of messages) {
            if (!msg?.id || seenMessageIds.has(msg.id)) continue;
            seenMessageIds.add(msg.id);

            try {
                const data = await gmailClient.users.messages.get({
                    userId: 'me',
                    id: msg.id,
                    format: 'metadata',
                    metadataHeaders: ['From', 'To', 'Cc', 'Reply-To']
                });

                const headers = data?.data?.payload?.headers || [];
                for (const header of headers) {
                    const value = header?.value || '';
                    const normalizedHeader = normalizeIdentity(value);
                    const matchedIdentity = keyTokens.length > 1
                        ? keyTokens.every(token => normalizedHeader.includes(token))
                        : normalizedHeader.includes(key);
                    const emails = extractEmailsFromText(value);

                    for (const email of emails) {
                        if (ownEmail && email === ownEmail) continue;
                        if (isDisallowedAttendeeEmail(email)) continue;
                        const localPart = normalizeIdentity(email.split('@')[0] || '');
                        const matchedLocalPart = keyTokens.length > 1
                            ? keyTokens.every(token => localPart.includes(token))
                            : localPart.includes(key);
                        if (!matchedIdentity && !matchedLocalPart) continue;
                        if (requiresStrictHeaderMatch && !matchedIdentity && !matchedLocalPart) continue;
                        const existing = scoreByEmail.get(email) || 0;
                        const bump = matchedIdentity && matchedLocalPart
                            ? 8
                            : (matchedIdentity ? 6 : 4);
                        scoreByEmail.set(email, existing + bump);
                    }
                }
            } catch {
                continue;
            }
        }
    }

    if (scoreByEmail.size === 0) return null;
    const ranked = [...scoreByEmail.entries()].sort((a, b) => b[1] - a[1]);

    if (ranked.length > 1) {
        const topScore = ranked[0][1];
        const secondScore = ranked[1][1];
        if (topScore <= secondScore + 1) {
            return null;
        }
    }

    return ranked[0][0];
}

async function normalizeEventAttendees(attendeesInput = []) {
    if (!Array.isArray(attendeesInput) || attendeesInput.length === 0) return [];

    const ownEmail = await getPrimaryEmailAddress();
    const resolved = [];
    const unresolved = [];
    let selfOnlyRequested = false;

    for (const item of attendeesInput) {
        const raw = String(item || '').trim();
        if (!raw) continue;

        if (isSelfOnlyAttendeeDirective(raw)) {
            selfOnlyRequested = true;
            continue;
        }

        if (isSelfReferenceAttendeeToken(raw)) {
            if (ownEmail && !isDisallowedAttendeeEmail(ownEmail)) {
                resolved.push(ownEmail.toLowerCase());
            }
            continue;
        }

        if (isValidEmail(raw) && !isDisallowedAttendeeEmail(raw)) {
            resolved.push(raw.toLowerCase());
            continue;
        }

        const found = await resolveEmailFromGmailHistory(raw);
        if (found) resolved.push(found.toLowerCase());
        else unresolved.push(raw);
    }

    const deduped = [...new Set(resolved)];
    if (selfOnlyRequested) {
        if (ownEmail && !isDisallowedAttendeeEmail(ownEmail)) {
            return [ownEmail.toLowerCase()];
        }
        return [];
    }
    if (unresolved.length > 0) {
        throw new Error(`Could not resolve attendee email(s) from Gmail history: ${unresolved.join(', ')}. Please provide exact email address(es).`);
    }
    return deduped;
}

function parseIsoDateTime(value, fieldName) {
    const parsed = new Date(value);
    if (!value || Number.isNaN(parsed.getTime())) {
        throw new Error(`${fieldName} must be a valid ISO 8601 datetime.`);
    }
    return parsed;
}

function normalizeDurationMinutes(value, fallback = 30) {
    const minutes = Number(value);
    if (!Number.isFinite(minutes) || minutes <= 0) return fallback;
    return Math.round(minutes);
}

function mergeBusyIntervals(busy = [], rangeStart, rangeEnd) {
    const normalized = [];

    for (const interval of busy) {
        const start = new Date(interval?.start);
        const end = new Date(interval?.end);
        if (Number.isNaN(start.getTime()) || Number.isNaN(end.getTime()) || end <= start) continue;
        if (end <= rangeStart || start >= rangeEnd) continue;

        const clippedStart = start < rangeStart ? rangeStart : start;
        const clippedEnd = end > rangeEnd ? rangeEnd : end;
        if (clippedEnd > clippedStart) {
            normalized.push({ start: clippedStart, end: clippedEnd });
        }
    }

    normalized.sort((a, b) => a.start - b.start);
    const merged = [];
    for (const interval of normalized) {
        const previous = merged[merged.length - 1];
        if (!previous || interval.start > previous.end) {
            merged.push({ ...interval });
            continue;
        }
        if (interval.end > previous.end) {
            previous.end = interval.end;
        }
    }
    return merged;
}

function calculateFreeSlots({ timeMin, timeMax, busy = [], durationMinutes = 30 }) {
    const rangeStart = parseIsoDateTime(timeMin, 'timeMin');
    const rangeEnd = parseIsoDateTime(timeMax, 'timeMax');
    if (rangeEnd <= rangeStart) {
        throw new Error('timeMax must be greater than timeMin.');
    }

    const minimumDurationMs = normalizeDurationMinutes(durationMinutes) * 60 * 1000;
    const mergedBusy = mergeBusyIntervals(busy, rangeStart, rangeEnd);

    const freeSlots = [];
    let cursor = rangeStart;

    for (const interval of mergedBusy) {
        if (interval.start > cursor) {
            const gapMs = interval.start.getTime() - cursor.getTime();
            if (gapMs >= minimumDurationMs) {
                freeSlots.push({
                    start: cursor.toISOString(),
                    end: interval.start.toISOString(),
                    durationMinutes: Math.round(gapMs / 60000)
                });
            }
        }
        if (interval.end > cursor) {
            cursor = interval.end;
        }
    }

    if (rangeEnd > cursor) {
        const gapMs = rangeEnd.getTime() - cursor.getTime();
        if (gapMs >= minimumDurationMs) {
            freeSlots.push({
                start: cursor.toISOString(),
                end: rangeEnd.toISOString(),
                durationMinutes: Math.round(gapMs / 60000)
            });
        }
    }

    return {
        mergedBusy: mergedBusy.map(interval => ({
            start: interval.start.toISOString(),
            end: interval.end.toISOString()
        })),
        freeSlots
    };
}

function freeBusyErrorMessage(calendarId, errors = []) {
    const reasonText = errors
        .map(error => String(error?.reason || error?.message || '').toLowerCase())
        .join(' ');
    if (/notfound|forbidden|insufficientpermissions|accessdenied/.test(reasonText)) {
        return `Cannot access availability for "${calendarId}". Ask them to share their calendar with at least "See free/busy" permission.`;
    }
    return `Could not retrieve availability for "${calendarId}".`;
}

async function resolveAvailabilityCalendarId({ person, email, calendarId }) {
    const explicitCalendarId = String(calendarId || '').trim();
    if (explicitCalendarId) return explicitCalendarId;

    const explicitEmail = String(email || '').trim();
    if (explicitEmail) {
        if (!isValidEmail(explicitEmail) || isDisallowedAttendeeEmail(explicitEmail)) {
            throw new Error(`Invalid attendee email: ${explicitEmail}`);
        }
        return explicitEmail.toLowerCase();
    }

    const personText = String(person || '').trim();
    if (!personText) {
        throw new Error('Provide one of: person, email, or calendarId.');
    }
    if (isValidEmail(personText) && !isDisallowedAttendeeEmail(personText)) {
        return personText.toLowerCase();
    }

    const resolved = await resolveEmailFromGmailHistory(personText);
    if (!resolved) {
        throw new Error(`Could not resolve "${personText}" from Gmail history. Please provide exact email address.`);
    }
    return resolved.toLowerCase();
}

async function queryFreeBusyCalendars({ timeMin, timeMax, calendarIds = [] }) {
    const rangeStart = parseIsoDateTime(timeMin, 'timeMin');
    const rangeEnd = parseIsoDateTime(timeMax, 'timeMax');
    if (rangeEnd <= rangeStart) {
        throw new Error('timeMax must be greater than timeMin.');
    }
    if (!Array.isArray(calendarIds) || calendarIds.length === 0) {
        throw new Error('calendarIds must include at least one calendar.');
    }

    const dedupedCalendarIds = [...new Set(
        calendarIds
            .map(id => String(id || '').trim())
            .filter(Boolean)
    )];
    if (dedupedCalendarIds.length === 0) {
        throw new Error('calendarIds must include at least one non-empty calendar ID.');
    }

    const response = await calendarClient.freebusy.query({
        requestBody: {
            timeMin: rangeStart.toISOString(),
            timeMax: rangeEnd.toISOString(),
            items: dedupedCalendarIds.map(id => ({ id }))
        }
    });

    const calendars = {};
    const errors = {};
    for (const [id, data] of Object.entries(response?.data?.calendars || {})) {
        calendars[id] = { busy: data?.busy || [] };
        if (Array.isArray(data?.errors) && data.errors.length > 0) {
            errors[id] = data.errors;
        }
    }

    return {
        timeMin: rangeStart.toISOString(),
        timeMax: rangeEnd.toISOString(),
        requestedCalendarIds: dedupedCalendarIds,
        calendars,
        errors
    };
}

// 1. List Events
async function listEvents({ calendarId = 'primary', maxResults = 10, timeMin, timeMax }) {
    if (!calendarClient) throw new Error('Calendar not authenticated');
    const limit = clampPagination(maxResults, { fallback: 10, max: 250 });
    const params = { calendarId, maxResults: limit, singleEvents: true, orderBy: 'startTime' };
    if (timeMin) params.timeMin = timeMin;
    else params.timeMin = new Date().toISOString();
    if (timeMax) params.timeMax = timeMax;
    const response = await calendarClient.events.list(params);
    const events = (response.data.items || []).map(e => ({
        id: e.id, summary: e.summary || '(no title)', status: e.status,
        start: e.start?.dateTime || e.start?.date, end: e.end?.dateTime || e.end?.date,
        location: e.location, description: e.description,
        attendees: (e.attendees || []).map(a => ({ email: a.email, responseStatus: a.responseStatus })),
        htmlLink: e.htmlLink,
        meetLink: getMeetLinkFromEvent(e)
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
        meetLink: getMeetLinkFromEvent(e),
        message: `Event: ${e.summary}`
    };
}

// 3. Create Event
async function createEvent({ calendarId = 'primary', summary, description, location, startDateTime, endDateTime, startDate, endDate, attendees, recurrence, timeZone, createMeetLink = false }) {
    if (!calendarClient) throw new Error('Calendar not authenticated');
    const hasTimedStart = !!startDateTime;
    const hasTimedEnd = !!endDateTime;
    const hasAllDayStart = !!startDate;
    const hasAllDayEnd = !!endDate;
    const hasTimedRange = hasTimedStart || hasTimedEnd;
    const hasAllDayRange = hasAllDayStart || hasAllDayEnd;

    if (hasTimedRange && hasAllDayRange) {
        throw new Error('Use either startDateTime/endDateTime OR startDate/endDate, not both.');
    }
    if (hasTimedRange && (!hasTimedStart || !hasTimedEnd)) {
        throw new Error('Timed events require both startDateTime and endDateTime.');
    }
    if (hasAllDayRange && (!hasAllDayStart || !hasAllDayEnd)) {
        throw new Error('All-day events require both startDate and endDate.');
    }
    if (!hasTimedRange && !hasAllDayRange) {
        throw new Error('Please provide a schedule: startDateTime/endDateTime for timed events or startDate/endDate for all-day events.');
    }

    const event = { summary };
    if (description) event.description = description;
    if (location) event.location = location;
    const defaultTz = Intl.DateTimeFormat().resolvedOptions().timeZone || 'UTC';
    if (startDateTime) event.start = { dateTime: startDateTime, timeZone: timeZone || defaultTz };
    else if (startDate) event.start = { date: startDate };
    if (endDateTime) event.end = { dateTime: endDateTime, timeZone: timeZone || defaultTz };
    else if (endDate) event.end = { date: endDate };
    if (attendees) {
        const resolvedAttendees = await normalizeEventAttendees(attendees);
        if (resolvedAttendees.length > 0) {
            event.attendees = resolvedAttendees.map(email => ({ email }));
        }
    }
    if (recurrence) event.recurrence = recurrence;
    if (createMeetLink) {
        event.conferenceData = {
            createRequest: {
                requestId: crypto.randomUUID(),
                conferenceSolutionKey: { type: 'hangoutsMeet' }
            }
        };
    }

    const insertParams = { calendarId, requestBody: event };
    if (createMeetLink) insertParams.conferenceDataVersion = 1;
    const result = await calendarClient.events.insert(insertParams);
    const meetLink = getMeetLinkFromEvent(result.data);
    return { success: true, eventId: result.data.id, htmlLink: result.data.htmlLink, meetLink, message: `Event "${summary}" created` };
}

// 3b. Create Google Meet event
async function createMeetEvent({ calendarId = 'primary', summary, description, startDateTime, endDateTime, attendees = [], timeZone }) {
    if (!calendarClient) throw new Error('Calendar not authenticated');
    if (!summary || !startDateTime || !endDateTime) {
        throw new Error('summary, startDateTime, and endDateTime are required to create a Meet event');
    }

    const effectiveTz = timeZone || Intl.DateTimeFormat().resolvedOptions().timeZone || 'UTC';
    const resolvedAttendees = await normalizeEventAttendees(attendees);
    const requestBody = {
        summary,
        description,
        start: { dateTime: startDateTime, timeZone: effectiveTz },
        end: { dateTime: endDateTime, timeZone: effectiveTz },
        attendees: resolvedAttendees.map(email => ({ email })),
        conferenceData: {
            createRequest: {
                requestId: crypto.randomUUID(),
                conferenceSolutionKey: { type: 'hangoutsMeet' }
            }
        }
    };

    const result = await calendarClient.events.insert({
        calendarId,
        conferenceDataVersion: 1,
        requestBody
    });
    const meetLink = getMeetLinkFromEvent(result.data);
    return {
        success: true,
        eventId: result.data.id,
        htmlLink: result.data.htmlLink,
        meetLink,
        message: `Meet event "${summary}" created`
    };
}

// 3c. Add Meet link to existing event
async function addMeetLinkToEvent({ calendarId = 'primary', eventId }) {
    if (!calendarClient) throw new Error('Calendar not authenticated');
    if (!eventId) throw new Error('eventId is required');

    const existing = (await calendarClient.events.get({ calendarId, eventId })).data;
    const alreadyLinked = getMeetLinkFromEvent(existing);
    if (alreadyLinked) {
        return {
            success: true,
            eventId,
            meetLink: alreadyLinked,
            message: 'This event already has a Google Meet link'
        };
    }

    existing.conferenceData = {
        createRequest: {
            requestId: crypto.randomUUID(),
            conferenceSolutionKey: { type: 'hangoutsMeet' }
        }
    };

    const result = await calendarClient.events.patch({
        calendarId,
        eventId,
        conferenceDataVersion: 1,
        requestBody: { conferenceData: existing.conferenceData }
    });
    const meetLink = getMeetLinkFromEvent(result.data);
    return {
        success: true,
        eventId,
        meetLink,
        message: 'Google Meet link added to event'
    };
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
    const result = await calendarClient.events.update({ calendarId, eventId, requestBody: existing, sendUpdates: 'all' });
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
    const result = await queryFreeBusyCalendars({ timeMin, timeMax, calendarIds });
    const checkedIds = result.requestedCalendarIds || [];
    const defaultedToPrimaryOnly = checkedIds.length === 1 && checkedIds[0] === 'primary';
    const warning = defaultedToPrimaryOnly
        ? 'Only your primary calendar was checked. This is not another person\'s availability.'
        : null;
    return {
        calendarIdsChecked: checkedIds,
        calendars: result.calendars,
        errors: Object.fromEntries(
            Object.entries(result.errors).map(([id, errs]) => [id, freeBusyErrorMessage(id, errs)])
        ),
        warning,
        message: warning ? `Free/busy info retrieved. ${warning}` : 'Free/busy info retrieved'
    };
}

// 9b. Check one person's availability
async function checkPersonAvailability({ person, email, calendarId, timeMin, timeMax, durationMinutes = 30 }) {
    if (!calendarClient) throw new Error('Calendar not authenticated');

    const resolvedCalendarId = await resolveAvailabilityCalendarId({ person, email, calendarId });
    const result = await queryFreeBusyCalendars({
        timeMin,
        timeMax,
        calendarIds: [resolvedCalendarId]
    });

    const availabilityError = result.errors[resolvedCalendarId];
    if (availabilityError) {
        throw new Error(freeBusyErrorMessage(resolvedCalendarId, availabilityError));
    }

    const busy = result.calendars[resolvedCalendarId]?.busy || [];
    const { mergedBusy, freeSlots } = calculateFreeSlots({
        timeMin: result.timeMin,
        timeMax: result.timeMax,
        busy,
        durationMinutes
    });

    return {
        calendarId: resolvedCalendarId,
        resolvedEmail: isValidEmail(resolvedCalendarId) ? resolvedCalendarId : null,
        requestedIdentity: person || email || calendarId || resolvedCalendarId,
        timeMin: result.timeMin,
        timeMax: result.timeMax,
        busy: mergedBusy,
        freeSlots,
        message: `Found ${freeSlots.length} free slot(s) for ${resolvedCalendarId}`
    };
}

// 9c. Find common free slots across people/calendars
async function findCommonFreeSlots({ timeMin, timeMax, durationMinutes = 30, people = [], calendarIds = [], includePrimary = true }) {
    if (!calendarClient) throw new Error('Calendar not authenticated');

    const requestedCalendarIds = Array.isArray(calendarIds) ? [...calendarIds] : [];
    const resolvedPeople = [];

    if (Array.isArray(people)) {
        for (const person of people) {
            const identity = String(person || '').trim();
            if (!identity) continue;
            const id = await resolveAvailabilityCalendarId({ person: identity });
            requestedCalendarIds.push(id);
            resolvedPeople.push({ input: identity, calendarId: id });
        }
    }

    if (includePrimary) {
        requestedCalendarIds.push('primary');
    }

    const result = await queryFreeBusyCalendars({
        timeMin,
        timeMax,
        calendarIds: requestedCalendarIds
    });

    const successfulCalendarIds = result.requestedCalendarIds.filter(id => !result.errors[id]);
    if (successfulCalendarIds.length === 0) {
        const firstError = result.requestedCalendarIds[0];
        throw new Error(
            firstError
                ? freeBusyErrorMessage(firstError, result.errors[firstError] || [])
                : 'No calendars were available for free/busy lookup.'
        );
    }

    const combinedBusy = [];
    for (const id of successfulCalendarIds) {
        for (const interval of result.calendars[id]?.busy || []) {
            combinedBusy.push(interval);
        }
    }

    const { mergedBusy, freeSlots } = calculateFreeSlots({
        timeMin: result.timeMin,
        timeMax: result.timeMax,
        busy: combinedBusy,
        durationMinutes
    });

    const errorMessages = Object.fromEntries(
        Object.entries(result.errors).map(([id, errs]) => [id, freeBusyErrorMessage(id, errs)])
    );

    return {
        timeMin: result.timeMin,
        timeMax: result.timeMax,
        durationMinutes: normalizeDurationMinutes(durationMinutes),
        resolvedPeople,
        calendarsChecked: successfulCalendarIds,
        unavailableCalendars: errorMessages,
        mergedBusy,
        freeSlots,
        message: `Found ${freeSlots.length} common free slot(s) across ${successfulCalendarIds.length} calendar(s)`
    };
}

// 10. List Recurring Event Instances
async function listRecurringInstances({ calendarId = 'primary', eventId, maxResults = 10, timeMin, timeMax }) {
    if (!calendarClient) throw new Error('Calendar not authenticated');
    const limit = clampPagination(maxResults, { fallback: 10, max: 250 });
    const params = { calendarId, eventId, maxResults: limit };
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
async function updateEventAttendees({ calendarId = 'primary', eventId, addAttendees = [], removeAttendees = [], attendees }) {
    if (!calendarClient) throw new Error('Calendar not authenticated');
    const existing = (await calendarClient.events.get({ calendarId, eventId })).data;
    // Accept 'attendees' as fallback for 'addAttendees' (AI sometimes uses wrong param name)
    const toAdd = addAttendees.length > 0 ? addAttendees : (attendees || []);
    let currentAttendees = existing.attendees || [];
    if (toAdd.length > 0) {
        const resolvedAdditions = await normalizeEventAttendees(toAdd);
        const existingEmails = new Set(currentAttendees.map(a => a.email));
        for (const email of resolvedAdditions) {
            if (!existingEmails.has(email)) currentAttendees.push({ email });
        }
    }
    if (removeAttendees.length > 0) {
        const resolvedRemovals = await normalizeEventAttendees(removeAttendees);
        const removeSet = new Set(resolvedRemovals.map(email => email.toLowerCase()));
        currentAttendees = currentAttendees.filter(a => !removeSet.has(String(a.email || '').toLowerCase()));
    }
    existing.attendees = currentAttendees;
    const result = await calendarClient.events.update({ calendarId, eventId, requestBody: existing, sendUpdates: 'all' });
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
//  GOOGLE CHAT TOOL IMPLEMENTATIONS
// ============================================================

function normalizeSpaceName(spaceId) {
    if (!spaceId) return '';
    return spaceId.startsWith('spaces/') ? spaceId : `spaces/${spaceId}`;
}

async function listChatSpaces({ maxResults = 20 }) {
    if (!gchatClient) throw new Error('Google Chat not authenticated');
    const limit = clampPagination(maxResults, { fallback: 20, max: 100 });
    const response = await gchatClient.spaces.list({ pageSize: limit });
    const spaces = (response.data.spaces || []).map(space => ({
        name: space.name,
        displayName: space.displayName || space.name,
        spaceType: space.spaceType,
        singleUserBotDm: !!space.singleUserBotDm,
        threaded: !!space.threaded
    }));
    return { spaces, message: `Found ${spaces.length} Google Chat spaces` };
}

async function sendChatMessage({ spaceId, text }) {
    if (!gchatClient) throw new Error('Google Chat not authenticated');
    if (!spaceId || !text) throw new Error('spaceId and text are required');

    const parent = normalizeSpaceName(spaceId);
    const response = await gchatClient.spaces.messages.create({
        parent,
        requestBody: { text }
    });
    return {
        success: true,
        space: parent,
        messageName: response.data.name,
        createTime: response.data.createTime,
        text: response.data.text || text,
        message: `Message sent to ${parent}`
    };
}

async function listChatMessages({ spaceId, maxResults = 20 }) {
    if (!gchatClient) throw new Error('Google Chat not authenticated');
    if (!spaceId) throw new Error('spaceId is required');

    const parent = normalizeSpaceName(spaceId);
    const limit = clampPagination(maxResults, { fallback: 20, max: 100 });
    const response = await gchatClient.spaces.messages.list({
        parent,
        pageSize: limit
    });
    const messages = (response.data.messages || []).map(msg => ({
        name: msg.name,
        text: msg.text || '',
        createTime: msg.createTime,
        sender: msg.sender?.displayName || msg.sender?.name || 'Unknown',
        thread: msg.thread?.name
    }));
    return { space: parent, messages, message: `Found ${messages.length} messages in ${parent}` };
}

// ============================================================
//  GOOGLE DRIVE TOOL IMPLEMENTATIONS
// ============================================================

function normalizeDriveFile(record) {
    return {
        id: record.id,
        name: record.name,
        mimeType: record.mimeType,
        owners: (record.owners || []).map(owner => owner.displayName || owner.emailAddress).filter(Boolean),
        modifiedTime: record.modifiedTime,
        createdTime: record.createdTime,
        size: record.size ? Number(record.size) : null,
        webViewLink: record.webViewLink,
        parents: record.parents || [],
        trashed: !!record.trashed
    };
}

function escapeDriveQueryLiteral(value) {
    return String(value || '')
        .replace(/\\/g, '\\\\')
        .replace(/'/g, "\\'");
}

function looksLikeDriveQuerySyntax(query) {
    const value = String(query || '').trim();
    if (!value) return false;
    return /(^|\s)(name|mimeType|modifiedTime|createdTime|trashed|sharedWithMe|owners|parents)\s|contains|=|!=|<=|>=|<|>|\band\b|\bor\b|\bnot\b|\(|\)|'/.test(value);
}

function buildDriveKeywordClause(query) {
    const escaped = escapeDriveQueryLiteral(query);
    if (!escaped) return '';
    return `(name contains '${escaped}' or fullText contains '${escaped}')`;
}

function buildDriveQueryClause(query) {
    const trimmed = String(query || '').trim();
    if (!trimmed) return '';
    if (looksLikeDriveQuerySyntax(trimmed)) return `(${trimmed})`;
    return buildDriveKeywordClause(trimmed);
}

const DRIVE_EXPORT_FORMATS = {
    'application/vnd.google-apps.document': {
        defaultFormat: 'docx',
        formats: {
            docx: { mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', extension: '.docx' },
            pdf: { mimeType: 'application/pdf', extension: '.pdf' },
            txt: { mimeType: 'text/plain', extension: '.txt' }
        }
    },
    'application/vnd.google-apps.spreadsheet': {
        defaultFormat: 'xlsx',
        formats: {
            xlsx: { mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', extension: '.xlsx' },
            pdf: { mimeType: 'application/pdf', extension: '.pdf' },
            csv: { mimeType: 'text/csv', extension: '.csv' }
        }
    },
    'application/vnd.google-apps.presentation': {
        defaultFormat: 'pptx',
        formats: {
            pptx: { mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation', extension: '.pptx' },
            pdf: { mimeType: 'application/pdf', extension: '.pdf' },
            txt: { mimeType: 'text/plain', extension: '.txt' }
        }
    },
    'application/vnd.google-apps.drawing': {
        defaultFormat: 'pdf',
        formats: {
            pdf: { mimeType: 'application/pdf', extension: '.pdf' },
            png: { mimeType: 'image/png', extension: '.png' },
            svg: { mimeType: 'image/svg+xml', extension: '.svg' }
        }
    }
};

function sanitizeDownloadFilename(name, fallback = 'download') {
    const value = String(name || '').trim();
    const sanitized = value
        .replace(/[\\/:*?"<>|]+/g, '_')
        .replace(/\s+/g, ' ')
        .trim();
    return sanitized || fallback;
}

function ensureFilenameExtension(filename, extension) {
    const safeName = sanitizeDownloadFilename(filename, 'download');
    const ext = String(extension || '').trim();
    if (!ext) return safeName;
    if (safeName.toLowerCase().endsWith(ext.toLowerCase())) return safeName;
    return `${safeName}${ext.startsWith('.') ? ext : `.${ext}`}`;
}

function resolveDriveExportFormat(mimeType, requestedFormat) {
    const profile = DRIVE_EXPORT_FORMATS[String(mimeType || '')];
    if (!profile) return null;

    const formatKey = String(requestedFormat || profile.defaultFormat).trim().toLowerCase() || profile.defaultFormat;
    const selected = profile.formats[formatKey] || profile.formats[profile.defaultFormat];
    if (!selected) return null;
    return {
        format: profile.formats[formatKey] ? formatKey : profile.defaultFormat,
        mimeType: selected.mimeType,
        extension: selected.extension
    };
}

function buildDriveDownloadUrl(fileId, { format, filename } = {}) {
    const encodedFileId = encodeURIComponent(String(fileId || '').trim());
    const params = new URLSearchParams();
    if (format) params.set('format', String(format));
    if (filename) params.set('filename', String(filename));
    const query = params.toString();
    return `/api/drive/download/${encodedFileId}${query ? `?${query}` : ''}`;
}

async function extractPdfTextViaDriveConversion({ fileId, originalName = 'document.pdf' }) {
    if (!driveClient) throw new Error('Google Drive not authenticated');
    let tempDocId = null;
    const safeName = String(originalName || 'document').replace(/\.[^.]+$/, '');
    try {
        const copyResponse = await driveClient.files.copy({
            fileId,
            supportsAllDrives: true,
            requestBody: {
                name: `${safeName} (temporary text extract)`,
                mimeType: 'application/vnd.google-apps.document'
            },
            fields: 'id,name'
        });
        tempDocId = copyResponse?.data?.id || null;
        if (!tempDocId) throw new Error('Could not create temporary conversion document.');

        await sleep(250);
        const exportResponse = await driveClient.files.export(
            { fileId: tempDocId, mimeType: 'text/plain' },
            { responseType: 'arraybuffer' }
        );
        const text = Buffer.from(exportResponse.data).toString('utf8');
        return normalizeReadableText(text, 120000).text;
    } finally {
        if (tempDocId) {
            driveClient.files.delete({ fileId: tempDocId, supportsAllDrives: true }).catch(() => { });
        }
    }
}

async function listDriveFiles({ query, pageSize = 100, orderBy = 'modifiedTime desc', includeTrashed = false }) {
    if (!driveClient) throw new Error('Google Drive not authenticated');
    const limit = clampPagination(pageSize, { fallback: 100, max: 200 });
    const qParts = [];
    if (!includeTrashed) qParts.push('trashed = false');
    const queryText = String(query || '').trim();
    const primaryClause = buildDriveQueryClause(queryText);
    if (primaryClause) qParts.push(primaryClause);

    let response;
    try {
        response = await driveClient.files.list({
            q: qParts.length > 0 ? qParts.join(' and ') : undefined,
            pageSize: limit,
            orderBy,
            supportsAllDrives: true,
            includeItemsFromAllDrives: true,
            corpora: 'allDrives',
            fields: 'nextPageToken, files(id,name,mimeType,owners(displayName,emailAddress),modifiedTime,createdTime,size,webViewLink,parents,trashed)'
        });
    } catch (error) {
        const shouldRetryWithKeyword = queryText && looksLikeDriveQuerySyntax(queryText);
        if (!shouldRetryWithKeyword) throw error;

        const retryQParts = [];
        if (!includeTrashed) retryQParts.push('trashed = false');
        retryQParts.push(buildDriveKeywordClause(queryText));
        response = await driveClient.files.list({
            q: retryQParts.join(' and '),
            pageSize: limit,
            orderBy,
            supportsAllDrives: true,
            includeItemsFromAllDrives: true,
            corpora: 'allDrives',
            fields: 'nextPageToken, files(id,name,mimeType,owners(displayName,emailAddress),modifiedTime,createdTime,size,webViewLink,parents,trashed)'
        });
    }

    const files = (response.data.files || []).map(normalizeDriveFile);
    return { files, nextPageToken: response.data.nextPageToken || null, message: `Found ${files.length} Drive file(s)` };
}

async function getDriveFile({ fileId }) {
    if (!driveClient) throw new Error('Google Drive not authenticated');
    if (!fileId) throw new Error('fileId is required');
    const response = await driveClient.files.get({
        fileId,
        fields: 'id,name,mimeType,description,owners(displayName,emailAddress),modifiedTime,createdTime,size,webViewLink,webContentLink,parents,trashed'
    });
    return { file: normalizeDriveFile(response.data), webContentLink: response.data.webContentLink || null, description: response.data.description || '', message: `Drive file: ${response.data.name}` };
}

async function createDriveFolder({ name, parentId }) {
    if (!driveClient) throw new Error('Google Drive not authenticated');
    if (!name) throw new Error('name is required');
    const requestBody = {
        name,
        mimeType: 'application/vnd.google-apps.folder'
    };
    if (parentId) requestBody.parents = [parentId];

    const response = await driveClient.files.create({
        requestBody,
        fields: 'id,name,mimeType,webViewLink,parents'
    });
    return {
        success: true,
        file: normalizeDriveFile(response.data),
        message: `Folder "${response.data.name}" created`
    };
}

async function createDriveFile({ name, content = '', mimeType = 'text/plain', parentId, localPath }) {
    if (!driveClient) throw new Error('Google Drive not authenticated');
    if (!name) throw new Error('name is required');

    const requestBody = { name };
    if (parentId) requestBody.parents = [parentId];

    let fileBody = content;
    let resolvedMimeType = mimeType;

    // If localPath is provided, read the file from disk (validated to uploads dir)
    if (localPath) {
        if (!validateLocalPath(localPath)) throw new Error('File path not allowed. Only files in uploads directory can be used.');
    }
    if (localPath && fs.existsSync(localPath)) {
        fileBody = fs.createReadStream(localPath);
        // If mimeType is default and we have a file, try to infer from extension
        if (mimeType === 'text/plain' || !mimeType) {
            const ext = path.extname(localPath).toLowerCase();
            const mimeTypes = {
                '.pdf': 'application/pdf',
                '.doc': 'application/msword',
                '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                '.xls': 'application/vnd.ms-excel',
                '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                '.ppt': 'application/vnd.ms-powerpoint',
                '.pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
                '.jpg': 'image/jpeg',
                '.jpeg': 'image/jpeg',
                '.png': 'image/png',
                '.gif': 'image/gif',
                '.txt': 'text/plain',
                '.csv': 'text/csv',
                '.json': 'application/json',
                '.zip': 'application/zip'
            };
            resolvedMimeType = mimeTypes[ext] || 'application/octet-stream';
        }
    }

    const response = await driveClient.files.create({
        requestBody,
        media: { mimeType: resolvedMimeType, body: fileBody },
        fields: 'id,name,mimeType,webViewLink,parents,size,modifiedTime'
    });
    return {
        success: true,
        file: normalizeDriveFile(response.data),
        message: `File "${response.data.name}" created`,
        uploadedFromLocal: !!localPath
    };
}

async function updateDriveFile({ fileId, content, name, mimeType = 'text/plain' }) {
    if (!driveClient) throw new Error('Google Drive not authenticated');
    if (!fileId) throw new Error('fileId is required');

    const requestBody = {};
    if (name) requestBody.name = name;

    const params = {
        fileId,
        requestBody,
        fields: 'id,name,mimeType,webViewLink,parents,size,modifiedTime'
    };
    if (content !== undefined) {
        params.media = { mimeType, body: content };
    }

    const response = await driveClient.files.update(params);
    return {
        success: true,
        file: normalizeDriveFile(response.data),
        message: `File "${response.data.name}" updated`
    };
}

async function deleteDriveFile({ fileId, permanent = false }) {
    if (!driveClient) throw new Error('Google Drive not authenticated');
    if (!fileId) throw new Error('fileId is required');

    if (permanent) {
        await driveClient.files.delete({ fileId });
        return { success: true, fileId, message: 'Drive file permanently deleted' };
    }

    await driveClient.files.update({
        fileId,
        requestBody: { trashed: true },
        fields: 'id,name,trashed'
    });
    return { success: true, fileId, message: 'Drive file moved to trash' };
}

async function copyDriveFile({ fileId, name, parentId }) {
    if (!driveClient) throw new Error('Google Drive not authenticated');
    if (!fileId) throw new Error('fileId is required');

    const requestBody = {};
    if (name) requestBody.name = name;
    if (parentId) requestBody.parents = [parentId];

    const response = await driveClient.files.copy({
        fileId,
        requestBody,
        fields: 'id,name,mimeType,webViewLink,parents,size,modifiedTime'
    });
    return { success: true, file: normalizeDriveFile(response.data), message: `Copied file to "${response.data.name}"` };
}

async function moveDriveFile({ fileId, newParentId }) {
    if (!driveClient) throw new Error('Google Drive not authenticated');
    if (!fileId || !newParentId) throw new Error('fileId and newParentId are required');

    const meta = await driveClient.files.get({ fileId, fields: 'id,name,parents' });
    const currentParents = meta.data.parents || [];

    const response = await driveClient.files.update({
        fileId,
        addParents: newParentId,
        removeParents: currentParents.join(','),
        fields: 'id,name,mimeType,webViewLink,parents,modifiedTime'
    });

    return {
        success: true,
        file: normalizeDriveFile(response.data),
        previousParents: currentParents,
        message: `Moved "${response.data.name}" to new parent`
    };
}

async function shareDriveFile({ fileId, emailAddress, role = 'reader', sendNotificationEmail = true }) {
    if (!driveClient) throw new Error('Google Drive not authenticated');
    if (!fileId || !emailAddress) throw new Error('fileId and emailAddress are required');

    const permission = await driveClient.permissions.create({
        fileId,
        sendNotificationEmail,
        requestBody: {
            type: 'user',
            role,
            emailAddress
        },
        fields: 'id,type,role,emailAddress'
    });
    return {
        success: true,
        permission: permission.data,
        message: `Shared file with ${emailAddress} as ${role}`
    };
}

async function downloadDriveFile({ fileId, format, filename }) {
    if (!driveClient) throw new Error('Google Drive not authenticated');
    if (!fileId) throw new Error('fileId is required');

    const meta = await driveClient.files.get({
        fileId,
        supportsAllDrives: true,
        fields: 'id,name,mimeType,size,webViewLink'
    });
    const sourceName = sanitizeDownloadFilename(meta.data.name || 'download');
    const mimeType = String(meta.data.mimeType || 'application/octet-stream');

    let resolvedFormat = 'raw';
    let exportMimeType = mimeType;
    let downloadName = sanitizeDownloadFilename(filename || sourceName, sourceName);
    if (mimeType.startsWith('application/vnd.google-apps.')) {
        const exportInfo = resolveDriveExportFormat(mimeType, format);
        if (!exportInfo) {
            throw new Error(`Google Workspace file type "${mimeType}" is not supported for export download.`);
        }
        resolvedFormat = exportInfo.format;
        exportMimeType = exportInfo.mimeType;
        downloadName = ensureFilenameExtension(downloadName, exportInfo.extension);
    }

    const downloadUrl = buildDriveDownloadUrl(fileId, {
        format: resolvedFormat === 'raw' ? undefined : resolvedFormat,
        filename: downloadName
    });

    return {
        fileId,
        name: meta.data.name,
        mimeType,
        size: meta.data.size ? Number(meta.data.size) : null,
        format: resolvedFormat,
        exportMimeType,
        downloadName,
        downloadUrl,
        webViewLink: meta.data.webViewLink || null,
        message: `Download ready for "${meta.data.name}". Use the provided link to save the file.`
    };
}

// Download Drive file to local disk (for use as email attachment or upload to other services)
async function downloadDriveFileToLocal({ fileId, format }) {
    if (!driveClient) throw new Error('Google Drive not authenticated');
    if (!fileId) throw new Error('fileId is required');

    const meta = await driveClient.files.get({
        fileId,
        supportsAllDrives: true,
        fields: 'id,name,mimeType,size'
    });
    const sourceName = sanitizeDownloadFilename(meta.data.name || 'download');
    const mimeType = String(meta.data.mimeType || 'application/octet-stream');

    let downloadStream;
    let downloadName = sourceName;

    if (mimeType.startsWith('application/vnd.google-apps.')) {
        // Google Workspace file — need to export
        const exportInfo = resolveDriveExportFormat(mimeType, format);
        if (!exportInfo) {
            throw new Error(`Google Workspace file type "${mimeType}" is not supported for export.`);
        }
        downloadName = ensureFilenameExtension(downloadName, exportInfo.extension);
        const res = await driveClient.files.export(
            { fileId, mimeType: exportInfo.mimeType },
            { responseType: 'stream' }
        );
        downloadStream = res.data;
    } else {
        // Regular file — direct download
        const res = await driveClient.files.get(
            { fileId, alt: 'media' },
            { responseType: 'stream' }
        );
        downloadStream = res.data;
    }

    // Save to uploads directory
    const uniqueSuffix = `${Date.now()}-${crypto.randomBytes(6).toString('hex')}`;
    const ext = path.extname(downloadName);
    const basename = path.basename(downloadName, ext);
    const localFilename = `${basename}-${uniqueSuffix}${ext}`;
    const localPath = path.join(UPLOADS_DIR, localFilename);

    await new Promise((resolve, reject) => {
        const writeStream = fs.createWriteStream(localPath);
        downloadStream.pipe(writeStream);
        writeStream.on('finish', resolve);
        writeStream.on('error', reject);
    });

    // Register in uploadedFiles map so it can be referenced
    const localFileId = crypto.randomBytes(16).toString('hex');
    const stats = fs.statSync(localPath);
    uploadedFiles.set(localFileId, {
        fileId: localFileId,
        path: localPath,
        originalName: downloadName,
        size: stats.size,
        mimeType: mimeType,
        uploadedAt: Date.now()
    });

    const downloadUrl = `http://localhost:${PORT}/api/download/${localFilename}`;

    return {
        success: true,
        fileId: localFileId,
        localPath,
        name: downloadName,
        size: stats.size,
        mimeType,
        downloadUrl,
        message: `File "${downloadName}" downloaded successfully. Click here to download to your computer: ${downloadUrl}`
    };
}

async function extractDriveFileText({ fileId, maxBytes = 40000 }) {
    if (!driveClient) throw new Error('Google Drive not authenticated');
    if (!fileId) throw new Error('fileId is required');

    const meta = await driveClient.files.get({
        fileId,
        supportsAllDrives: true,
        fields: 'id,name,mimeType,size,webViewLink'
    });

    const mimeType = String(meta.data.mimeType || '');
    if (mimeType === 'application/vnd.google-apps.spreadsheet') {
        throw new Error('This file is a Google Sheet. Use Sheets tools (read_sheet_values, get_spreadsheet) to access data.');
    }

    const limit = clampPagination(maxBytes, { fallback: 40000, min: 1024, max: 120000 });
    const lowerMime = mimeType.toLowerCase();
    let content = '';
    let extractionMethod = 'raw';
    let rawByteLength = Number(meta.data.size) || 0;

    if (lowerMime === 'application/pdf') {
        try {
            const convertedText = await extractPdfTextViaDriveConversion({
                fileId,
                originalName: meta.data.name
            });
            const normalized = normalizeReadableText(convertedText, limit);
            content = normalized.text;
            extractionMethod = 'drive_pdf_to_text';
        } catch {
            const fileResponse = await driveClient.files.get(
                { fileId, alt: 'media', supportsAllDrives: true },
                { responseType: 'arraybuffer' }
            );
            const rawBytes = Buffer.from(fileResponse.data);
            rawByteLength = rawBytes.length;
            const printable = extractPrintableStringsFromBuffer(rawBytes, { maxChars: limit });
            if (printable) {
                content = printable;
                extractionMethod = 'pdf_printable_strings';
            } else {
                throw new Error('Could not extract readable text from this PDF. Try opening it in Drive and copying text, or convert it to Google Docs.');
            }
        }
    } else if (mimeType.startsWith('application/vnd.google-apps.')) {
        const exportResponse = await driveClient.files.export(
            { fileId, mimeType: 'text/plain' },
            { responseType: 'arraybuffer' }
        );
        const normalized = normalizeReadableText(Buffer.from(exportResponse.data).toString('utf8'), limit);
        content = normalized.text;
        extractionMethod = 'google_workspace_export';
    } else {
        const fileResponse = await driveClient.files.get(
            { fileId, alt: 'media', supportsAllDrives: true },
            { responseType: 'arraybuffer' }
        );
        const rawBytes = Buffer.from(fileResponse.data);
        rawByteLength = rawBytes.length;

        if (isLikelyTextMimeType(mimeType) || !hasHighBinaryRatio(rawBytes)) {
            const normalized = normalizeReadableText(rawBytes.toString('utf8'), limit);
            content = normalized.text;
            extractionMethod = 'plain_text';
        } else {
            const printable = extractPrintableStringsFromBuffer(rawBytes, { maxChars: limit });
            if (!printable) {
                throw new Error(`"${meta.data.name}" is a binary file (${mimeType}). Text extraction is not available for this format.`);
            }
            content = printable;
            extractionMethod = 'binary_printable_strings';
        }
    }

    const normalizedContent = normalizeReadableText(content, limit);
    content = normalizedContent.text;
    if (!content) {
        throw new Error(`No readable text found in "${meta.data.name}".`);
    }

    const returnedBytes = Buffer.byteLength(content, 'utf8');
    const truncated = normalizedContent.truncated || returnedBytes >= limit;

    return {
        fileId,
        name: meta.data.name,
        mimeType,
        content,
        extractionMethod,
        byteLength: rawByteLength || returnedBytes,
        returnedBytes,
        truncated,
        webViewLink: meta.data.webViewLink || null,
        message: truncated
            ? `Downloaded truncated readable content for "${meta.data.name}" using ${extractionMethod} (${returnedBytes}/${rawByteLength || returnedBytes} bytes).`
            : `Downloaded readable content for "${meta.data.name}" using ${extractionMethod}.`
    };
}

async function appendDriveDocumentText({ fileId, text }) {
    if (!driveClient) throw new Error('Google Drive not authenticated');
    if (!fileId) throw new Error('fileId is required');
    const appendTextValue = String(text || '').replace(/\r\n?/g, '\n').replace(/\u0000/g, '').trim();
    if (!appendTextValue) throw new Error('text is required');

    const metaResponse = await driveClient.files.get({
        fileId,
        supportsAllDrives: true,
        fields: 'id,name,mimeType,webViewLink,size'
    });
    const file = metaResponse.data || {};
    const mimeType = String(file.mimeType || '');
    if (mimeType === 'application/vnd.google-apps.folder') {
        throw new Error('Cannot append text to a folder.');
    }
    if (mimeType === 'application/vnd.google-apps.spreadsheet') {
        throw new Error('Target is a spreadsheet. Use Sheets tools for spreadsheet updates.');
    }

    if (mimeType === 'application/vnd.google-apps.document' && docsClient && hasDocsScope()) {
        const prefix = appendTextValue.startsWith('\n') ? '' : '\n';
        await appendText({ documentId: fileId, text: `${prefix}${appendTextValue}` });
        return {
            success: true,
            fileId,
            name: file.name,
            mimeType,
            appendedText: appendTextValue,
            usedDocsApi: true,
            webViewLink: file.webViewLink || null,
            message: `Appended text to "${file.name}" using Google Docs API.`
        };
    }

    let currentContent = '';
    if (mimeType.startsWith('application/vnd.google-apps.')) {
        const exportResponse = await driveClient.files.export(
            { fileId, mimeType: 'text/plain' },
            { responseType: 'arraybuffer' }
        );
        currentContent = Buffer.from(exportResponse.data).toString('utf8');
    } else {
        const fileResponse = await driveClient.files.get(
            { fileId, alt: 'media', supportsAllDrives: true },
            { responseType: 'arraybuffer' }
        );
        const rawBytes = Buffer.from(fileResponse.data);
        if (!isLikelyTextMimeType(mimeType) && hasHighBinaryRatio(rawBytes)) {
            throw new Error(`"${file.name}" appears to be binary (${mimeType || 'unknown'}). Text append is not supported.`);
        }
        currentContent = rawBytes.toString('utf8');
    }

    currentContent = String(currentContent || '').replace(/\r\n?/g, '\n').replace(/\u0000/g, '');
    if (currentContent.length > DRIVE_TEXT_EDIT_MAX_CHARS) {
        throw new Error(`Document is too large to safely rewrite via Drive fallback (${currentContent.length} chars). Reconnect Docs and retry.`);
    }

    const base = currentContent.trimEnd();
    const updatedContent = base ? `${base}\n${appendTextValue}\n` : `${appendTextValue}\n`;
    const updated = await updateDriveFile({
        fileId,
        content: updatedContent,
        mimeType: 'text/plain'
    });

    return {
        success: true,
        fileId,
        name: updated.file?.name || file.name,
        mimeType,
        appendedText: appendTextValue,
        previousChars: currentContent.length,
        newChars: updatedContent.length,
        usedDocsApi: false,
        webViewLink: updated.file?.webViewLink || file.webViewLink || null,
        message: `Appended text to "${updated.file?.name || file.name}" using Drive fallback.`
    };
}

async function convertFileToGoogleDoc({ fileId, title, parentId, downloadConverted = false, downloadFormat = 'docx' }) {
    if (!driveClient) throw new Error('Google Drive not authenticated');
    if (!fileId) throw new Error('fileId is required');

    const source = await driveClient.files.get({
        fileId,
        supportsAllDrives: true,
        fields: 'id,name,mimeType,webViewLink,size'
    });
    const sourceMimeType = String(source.data.mimeType || '');
    if (sourceMimeType === 'application/vnd.google-apps.folder') {
        throw new Error('Cannot convert a folder to Google Docs.');
    }

    const desiredName = sanitizeDownloadFilename(
        title || `${String(source.data.name || 'Converted Document').replace(/\.[^.]+$/, '')} (Converted)`
    );
    const requestBody = {
        name: desiredName,
        mimeType: 'application/vnd.google-apps.document'
    };
    if (parentId) requestBody.parents = [parentId];

    const converted = await driveClient.files.copy({
        fileId,
        supportsAllDrives: true,
        requestBody,
        fields: 'id,name,mimeType,webViewLink,parents,modifiedTime'
    });

    const result = {
        success: true,
        sourceFileId: fileId,
        sourceName: source.data.name,
        sourceMimeType,
        sourceWebViewLink: source.data.webViewLink || null,
        documentId: converted.data.id,
        name: converted.data.name,
        mimeType: converted.data.mimeType,
        webViewLink: converted.data.webViewLink || null,
        parents: converted.data.parents || [],
        modifiedTime: converted.data.modifiedTime,
        message: `Converted "${source.data.name}" to Google Doc "${converted.data.name}".`
    };

    if (downloadConverted) {
        const exportInfo = resolveDriveExportFormat('application/vnd.google-apps.document', downloadFormat);
        if (!exportInfo) {
            throw new Error(`Unsupported download format "${downloadFormat}" for converted Google Doc.`);
        }
        const downloadName = ensureFilenameExtension(converted.data.name, exportInfo.extension);
        result.downloadFormat = exportInfo.format;
        result.downloadName = downloadName;
        result.downloadUrl = buildDriveDownloadUrl(converted.data.id, {
            format: exportInfo.format,
            filename: downloadName
        });
    }

    return result;
}

async function convertFileToGoogleSheet({ fileId, title, parentId, downloadConverted = false, downloadFormat = 'xlsx' }) {
    if (!driveClient) throw new Error('Google Drive not authenticated');
    if (!fileId) throw new Error('fileId is required');

    const source = await driveClient.files.get({
        fileId,
        supportsAllDrives: true,
        fields: 'id,name,mimeType,webViewLink,size'
    });
    const sourceMimeType = String(source.data.mimeType || '');
    if (sourceMimeType === 'application/vnd.google-apps.folder') {
        throw new Error('Cannot convert a folder to Google Sheets.');
    }

    const desiredName = sanitizeDownloadFilename(
        title || `${String(source.data.name || 'Converted Spreadsheet').replace(/\.[^.]+$/, '')} (Converted)`
    );
    const requestBody = {
        name: desiredName,
        mimeType: 'application/vnd.google-apps.spreadsheet'
    };
    if (parentId) requestBody.parents = [parentId];

    let converted;
    try {
        converted = await driveClient.files.copy({
            fileId,
            supportsAllDrives: true,
            requestBody,
            fields: 'id,name,mimeType,webViewLink,parents,modifiedTime'
        });
    } catch (error) {
        const message = String(error?.message || '').toLowerCase();
        if (message.includes('convert') || message.includes('mime')) {
            throw new Error('Google Drive could not convert this file into a Google Sheet. Try CSV/XLSX input or convert to a Google Doc instead.');
        }
        throw error;
    }

    const result = {
        success: true,
        sourceFileId: fileId,
        sourceName: source.data.name,
        sourceMimeType,
        sourceWebViewLink: source.data.webViewLink || null,
        spreadsheetId: converted.data.id,
        name: converted.data.name,
        mimeType: converted.data.mimeType,
        webViewLink: converted.data.webViewLink || null,
        parents: converted.data.parents || [],
        modifiedTime: converted.data.modifiedTime,
        message: `Converted "${source.data.name}" to Google Sheet "${converted.data.name}".`
    };

    if (downloadConverted) {
        const exportInfo = resolveDriveExportFormat('application/vnd.google-apps.spreadsheet', downloadFormat);
        if (!exportInfo) {
            throw new Error(`Unsupported download format "${downloadFormat}" for converted Google Sheet.`);
        }
        const downloadName = ensureFilenameExtension(converted.data.name, exportInfo.extension);
        result.downloadFormat = exportInfo.format;
        result.downloadName = downloadName;
        result.downloadUrl = buildDriveDownloadUrl(converted.data.id, {
            format: exportInfo.format,
            filename: downloadName
        });
    }

    return result;
}

// ============================================================
//  GOOGLE SHEETS TOOL IMPLEMENTATIONS
// ============================================================

function normalizeSheetValuesInput(values) {
    if (!Array.isArray(values)) {
        throw new Error('values must be an array');
    }
    if (values.length === 0) return [];
    if (Array.isArray(values[0])) return values;
    return [values];
}

function sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}

function normalizeSheetCellForCompare(value) {
    if (value === null || value === undefined) return '';
    if (typeof value === 'string') return value.trim();
    if (typeof value === 'number' && Number.isFinite(value)) return String(value);
    if (typeof value === 'boolean') return value ? 'true' : 'false';
    return String(value).trim();
}

function columnLettersToIndex(columnLetters) {
    const text = String(columnLetters || '').trim().toUpperCase();
    if (!/^[A-Z]+$/.test(text)) throw new Error(`Invalid column letters: ${columnLetters}`);
    let value = 0;
    for (let i = 0; i < text.length; i += 1) {
        value = value * 26 + (text.charCodeAt(i) - 64);
    }
    return value - 1;
}

function columnIndexToLetters(columnIndex) {
    const index = Number(columnIndex);
    if (!Number.isInteger(index) || index < 0) {
        throw new Error(`Invalid column index: ${columnIndex}`);
    }
    let current = index + 1;
    let letters = '';
    while (current > 0) {
        const rem = (current - 1) % 26;
        letters = String.fromCharCode(65 + rem) + letters;
        current = Math.floor((current - 1) / 26);
    }
    return letters;
}

function parseA1StartRow(a1Range) {
    const text = String(a1Range || '');
    const bangIndex = text.indexOf('!');
    const local = bangIndex >= 0 ? text.slice(bangIndex + 1) : text;
    const startRef = local.split(':')[0] || local;
    const match = startRef.match(/\d+/);
    return match ? Number(match[0]) : 1;
}

function isEffectivelyEmptyTimesheetCell(value) {
    const text = normalizeSheetCellForCompare(value).toLowerCase();
    return text === '' || text === 'na' || text === 'n/a' || text === '-' || text === 'none' || text === 'null';
}

function tryParseNumber(value) {
    if (value === null || value === undefined || value === '') return null;
    if (typeof value === 'number' && Number.isFinite(value)) return value;
    const normalized = String(value).trim();
    if (!/^[-+]?\d+(\.\d+)?$/.test(normalized)) return null;
    const parsed = Number(normalized);
    return Number.isFinite(parsed) ? parsed : null;
}

const MONTH_NAME_TO_NUMBER = {
    jan: 1, january: 1,
    feb: 2, february: 2,
    mar: 3, march: 3,
    apr: 4, april: 4,
    may: 5,
    jun: 6, june: 6,
    jul: 7, july: 7,
    aug: 8, august: 8,
    sep: 9, sept: 9, september: 9,
    oct: 10, october: 10,
    nov: 11, november: 11,
    dec: 12, december: 12
};

function toDateKey(year, month, day) {
    const y = Number(year);
    const m = Number(month);
    const d = Number(day);
    if (!Number.isInteger(y) || !Number.isInteger(m) || !Number.isInteger(d)) return null;
    if (y < 1900 || y > 9999 || m < 1 || m > 12 || d < 1 || d > 31) return null;
    const dt = new Date(y, m - 1, d);
    if (Number.isNaN(dt.getTime())) return null;
    if (dt.getFullYear() !== y || dt.getMonth() + 1 !== m || dt.getDate() !== d) return null;
    return `${String(y).padStart(4, '0')}-${String(m).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
}

function sheetSerialToDateKey(serial) {
    if (!Number.isFinite(serial)) return null;
    const wholeDays = Math.floor(serial);
    if (wholeDays < 1 || wholeDays > 300000) return null;
    const epochMs = Date.UTC(1899, 11, 30);
    const ms = epochMs + wholeDays * 86400000;
    const d = new Date(ms);
    if (Number.isNaN(d.getTime())) return null;
    const y = d.getUTCFullYear();
    const m = String(d.getUTCMonth() + 1).padStart(2, '0');
    const day = String(d.getUTCDate()).padStart(2, '0');
    return `${y}-${m}-${day}`;
}

function tryParseDateKey(value) {
    if (value === null || value === undefined || value === '') return null;
    if (value instanceof Date) {
        if (Number.isNaN(value.getTime())) return null;
        return toDateKey(value.getFullYear(), value.getMonth() + 1, value.getDate());
    }
    if (typeof value === 'number') return sheetSerialToDateKey(value);

    const text = String(value).trim();
    if (!text) return null;

    const cleaned = text
        .toLowerCase()
        .replace(/,/g, ' ')
        .replace(/\b(\d{1,2})(st|nd|rd|th)\b/g, '$1')
        .replace(/\s+/g, ' ')
        .trim();

    let match = cleaned.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})$/);
    if (match) {
        return toDateKey(match[1], match[2], match[3]);
    }

    match = cleaned.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})$/);
    if (match) {
        const first = Number(match[1]);
        const second = Number(match[2]);
        const year = match[3];
        if (first > 12 && second <= 12) return toDateKey(year, second, first);
        if (second > 12 && first <= 12) return toDateKey(year, first, second);
        return toDateKey(year, first, second);
    }

    match = cleaned.match(/^(\d{1,2})[\s\-\/]+([a-z]+)[\s\-\/]+(\d{4})$/);
    if (match) {
        const month = MONTH_NAME_TO_NUMBER[match[2]];
        if (month) return toDateKey(match[3], month, match[1]);
    }

    match = cleaned.match(/^([a-z]+)[\s\-\/]+(\d{1,2})[\s\-\/]+(\d{4})$/);
    if (match) {
        const month = MONTH_NAME_TO_NUMBER[match[1]];
        if (month) return toDateKey(match[3], month, match[2]);
    }

    const parsed = new Date(text);
    if (Number.isNaN(parsed.getTime())) return null;
    return toDateKey(parsed.getFullYear(), parsed.getMonth() + 1, parsed.getDate());
}

function sheetCellsEquivalent(expectedValue, actualValue) {
    const expected = normalizeSheetCellForCompare(expectedValue);
    const actual = normalizeSheetCellForCompare(actualValue);
    if (expected === actual) return true;

    const expectedNum = tryParseNumber(expectedValue);
    const actualNum = tryParseNumber(actualValue);
    if (expectedNum !== null && actualNum !== null && Math.abs(expectedNum - actualNum) < 1e-9) {
        return true;
    }

    const expectedDate = tryParseDateKey(expectedValue);
    const actualDate = tryParseDateKey(actualValue);
    if (expectedDate && actualDate && expectedDate === actualDate) {
        return true;
    }

    return false;
}

function detectFormulaInValues(values) {
    return values.some(row =>
        Array.isArray(row) && row.some(cell => typeof cell === 'string' && cell.trim().startsWith('='))
    );
}

function sheetValuesMatch(expected, actual) {
    for (let rowIndex = 0; rowIndex < expected.length; rowIndex += 1) {
        const expectedRow = Array.isArray(expected[rowIndex]) ? expected[rowIndex] : [];
        const actualRow = Array.isArray(actual[rowIndex]) ? actual[rowIndex] : [];
        for (let colIndex = 0; colIndex < expectedRow.length; colIndex += 1) {
            const expectedCell = expectedRow[colIndex];
            const actualCell = actualRow[colIndex];
            if (!sheetCellsEquivalent(expectedCell, actualCell)) return false;
        }
    }
    return true;
}

async function verifySheetWriteWithRetry({
    spreadsheetId,
    range,
    expectedValues,
    isFormulaWrite = false,
    maxAttempts = 3
}) {
    const renderOptions = isFormulaWrite
        ? ['FORMULA']
        : ['UNFORMATTED_VALUE', 'FORMATTED_VALUE'];

    let lastReadBackValues = [];
    let lastRenderOption = renderOptions[0];

    for (let attempt = 1; attempt <= maxAttempts; attempt += 1) {
        for (const renderOption of renderOptions) {
            const verifyResponse = await sheetsClient.spreadsheets.values.get({
                spreadsheetId,
                range,
                valueRenderOption: renderOption
            });
            const readBackValues = verifyResponse.data.values || [];
            lastReadBackValues = readBackValues;
            lastRenderOption = renderOption;
            if (sheetValuesMatch(expectedValues, readBackValues)) {
                return {
                    verified: true,
                    readBackValues,
                    renderOption,
                    attempts: attempt
                };
            }
        }
        if (attempt < maxAttempts) {
            await sleep(150 * attempt);
        }
    }

    return {
        verified: false,
        readBackValues: lastReadBackValues,
        renderOption: lastRenderOption,
        attempts: maxAttempts
    };
}

async function listSpreadsheets({ query, maxResults = 100 }) {
    if (!driveClient) throw new Error('Google Drive is required to list spreadsheets. Reconnect Google Drive.');
    const limit = clampPagination(maxResults, { fallback: 100, max: 200 });

    const qParts = [
        "mimeType = 'application/vnd.google-apps.spreadsheet'",
        'trashed = false'
    ];
    const queryText = String(query || '').trim();
    const primaryClause = buildDriveQueryClause(queryText);
    if (primaryClause) qParts.push(primaryClause);

    let response;
    try {
        response = await driveClient.files.list({
            q: qParts.join(' and '),
            pageSize: limit,
            orderBy: 'modifiedTime desc',
            supportsAllDrives: true,
            includeItemsFromAllDrives: true,
            corpora: 'allDrives',
            fields: 'files(id,name,owners(displayName,emailAddress),modifiedTime,webViewLink)'
        });
    } catch (error) {
        const shouldRetryWithKeyword = queryText && looksLikeDriveQuerySyntax(queryText);
        if (!shouldRetryWithKeyword) throw error;

        const retryQParts = [
            "mimeType = 'application/vnd.google-apps.spreadsheet'",
            'trashed = false',
            buildDriveKeywordClause(queryText)
        ];
        response = await driveClient.files.list({
            q: retryQParts.join(' and '),
            pageSize: limit,
            orderBy: 'modifiedTime desc',
            supportsAllDrives: true,
            includeItemsFromAllDrives: true,
            corpora: 'allDrives',
            fields: 'files(id,name,owners(displayName,emailAddress),modifiedTime,webViewLink)'
        });
    }

    let files = response.data.files || [];
    if (queryText && files.length === 0) {
        const broad = await driveClient.files.list({
            q: "mimeType = 'application/vnd.google-apps.spreadsheet' and trashed = false",
            pageSize: Math.max(limit, 200),
            orderBy: 'modifiedTime desc',
            supportsAllDrives: true,
            includeItemsFromAllDrives: true,
            corpora: 'allDrives',
            fields: 'files(id,name,owners(displayName,emailAddress),modifiedTime,webViewLink)'
        });
        const qLower = queryText.toLowerCase();
        files = (broad.data.files || []).filter(file => String(file.name || '').toLowerCase().includes(qLower));
    }

    const spreadsheets = files.map(file => ({
        spreadsheetId: file.id,
        title: file.name,
        owners: (file.owners || []).map(owner => owner.displayName || owner.emailAddress).filter(Boolean),
        modifiedTime: file.modifiedTime,
        webViewLink: file.webViewLink
    }));
    return { spreadsheets, message: `Found ${spreadsheets.length} spreadsheet(s)` };
}

async function createSpreadsheet({ title, sheets = [] }) {
    if (!sheetsClient) throw new Error('Google Sheets not authenticated');
    if (!title) throw new Error('title is required');

    const normalizedSheets = Array.isArray(sheets)
        ? sheets
            .map(name => String(name || '').trim())
            .filter(Boolean)
            .map(name => ({ properties: { title: name } }))
        : [];

    const response = await sheetsClient.spreadsheets.create({
        requestBody: {
            properties: { title },
            sheets: normalizedSheets.length > 0 ? normalizedSheets : undefined
        }
    });

    return {
        success: true,
        spreadsheetId: response.data.spreadsheetId,
        title: response.data.properties?.title || title,
        url: response.data.spreadsheetUrl,
        message: `Spreadsheet "${title}" created`
    };
}

async function getSpreadsheet({ spreadsheetId, includeGridData = false }) {
    if (!sheetsClient) throw new Error('Google Sheets not authenticated');
    if (!spreadsheetId) throw new Error('spreadsheetId is required');

    const response = await sheetsClient.spreadsheets.get({
        spreadsheetId,
        includeGridData
    });
    const data = response.data;
    return {
        spreadsheetId: data.spreadsheetId,
        title: data.properties?.title,
        locale: data.properties?.locale,
        timeZone: data.properties?.timeZone,
        spreadsheetUrl: data.spreadsheetUrl,
        sheetCount: (data.sheets || []).length,
        sheets: (data.sheets || []).map(sheet => ({
            sheetId: sheet.properties?.sheetId,
            title: sheet.properties?.title,
            index: sheet.properties?.index,
            rowCount: sheet.properties?.gridProperties?.rowCount,
            columnCount: sheet.properties?.gridProperties?.columnCount
        })),
        message: `Spreadsheet: ${data.properties?.title || spreadsheetId}`
    };
}

async function listSheetTabs({ spreadsheetId }) {
    if (!sheetsClient) throw new Error('Google Sheets not authenticated');
    if (!spreadsheetId) throw new Error('spreadsheetId is required');

    const response = await sheetsClient.spreadsheets.get({
        spreadsheetId,
        fields: 'spreadsheetId,properties.title,sheets.properties(sheetId,title,index,gridProperties(rowCount,columnCount))'
    });
    const tabs = (response.data.sheets || []).map(sheet => ({
        sheetId: sheet.properties?.sheetId,
        title: sheet.properties?.title,
        index: sheet.properties?.index,
        rowCount: sheet.properties?.gridProperties?.rowCount,
        columnCount: sheet.properties?.gridProperties?.columnCount
    }));
    return { spreadsheetId, title: response.data.properties?.title, tabs, message: `Found ${tabs.length} sheet tab(s)` };
}

async function addSheetTab({ spreadsheetId, title, rows = 1000, columns = 26 }) {
    if (!sheetsClient) throw new Error('Google Sheets not authenticated');
    if (!spreadsheetId || !title) throw new Error('spreadsheetId and title are required');

    const response = await sheetsClient.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: {
            requests: [{
                addSheet: {
                    properties: {
                        title,
                        gridProperties: {
                            rowCount: rows,
                            columnCount: columns
                        }
                    }
                }
            }]
        }
    });

    const sheetProperties = response.data.replies?.[0]?.addSheet?.properties || {};
    return {
        success: true,
        sheetId: sheetProperties.sheetId,
        title: sheetProperties.title || title,
        message: `Sheet tab "${title}" added`
    };
}

async function deleteSheetTab({ spreadsheetId, sheetId }) {
    if (!sheetsClient) throw new Error('Google Sheets not authenticated');
    if (!spreadsheetId || sheetId === undefined || sheetId === null) {
        throw new Error('spreadsheetId and sheetId are required');
    }

    await sheetsClient.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: {
            requests: [{
                deleteSheet: {
                    sheetId: Number(sheetId)
                }
            }]
        }
    });

    return { success: true, spreadsheetId, sheetId: Number(sheetId), message: `Deleted sheet tab ${sheetId}` };
}

async function readSheetValues({ spreadsheetId, range, valueRenderOption = 'FORMATTED_VALUE', dateTimeRenderOption = 'SERIAL_NUMBER' }) {
    if (!sheetsClient) throw new Error('Google Sheets not authenticated');
    if (!spreadsheetId || !range) throw new Error('spreadsheetId and range are required');

    const response = await sheetsClient.spreadsheets.values.get({
        spreadsheetId,
        range,
        valueRenderOption,
        dateTimeRenderOption
    });

    const values = response.data.values || [];
    return {
        spreadsheetId,
        range: response.data.range || range,
        majorDimension: response.data.majorDimension || 'ROWS',
        values,
        rowCount: values.length,
        message: `Read ${values.length} row(s) from ${range}`
    };
}

async function updateSheetValues({ spreadsheetId, range, values, valueInputOption = 'USER_ENTERED', majorDimension = 'ROWS' }) {
    if (!sheetsClient) throw new Error('Google Sheets not authenticated');
    if (!spreadsheetId || !range) throw new Error('spreadsheetId and range are required');

    const normalizedValues = normalizeSheetValuesInput(values);
    const response = await sheetsClient.spreadsheets.values.update({
        spreadsheetId,
        range,
        valueInputOption,
        requestBody: {
            majorDimension,
            values: normalizedValues
        }
    });

    const updatedCells = response.data.updatedCells || 0;
    if (updatedCells === 0) {
        throw new Error(`No cells were updated for ${range}. Verify sheet tab/range and try again.`);
    }

    const verification = await verifySheetWriteWithRetry({
        spreadsheetId,
        range,
        expectedValues: normalizedValues,
        isFormulaWrite: detectFormulaInValues(normalizedValues)
    });
    if (!verification.verified) {
        throw new Error(`Sheets write verification failed for ${range}. Read-back values did not match the requested update.`);
    }

    return {
        success: true,
        spreadsheetId,
        range,
        updatedRows: response.data.updatedRows || 0,
        updatedColumns: response.data.updatedColumns || 0,
        updatedCells,
        verified: true,
        readBackValues: verification.readBackValues,
        verificationRenderOption: verification.renderOption,
        verificationAttempts: verification.attempts,
        message: `Updated and verified ${updatedCells} cell(s)`
    };
}

async function appendSheetValues({ spreadsheetId, range, values, valueInputOption = 'USER_ENTERED', insertDataOption = 'INSERT_ROWS' }) {
    if (!sheetsClient) throw new Error('Google Sheets not authenticated');
    if (!spreadsheetId || !range) throw new Error('spreadsheetId and range are required');

    const normalizedValues = normalizeSheetValuesInput(values);
    const response = await sheetsClient.spreadsheets.values.append({
        spreadsheetId,
        range,
        valueInputOption,
        insertDataOption,
        requestBody: {
            values: normalizedValues
        }
    });

    const updates = response.data.updates || {};
    const updatedCells = updates.updatedCells || 0;
    if (updatedCells === 0) {
        throw new Error(`No rows were appended for ${range}. Verify target range and try again.`);
    }

    const updatedRange = updates.updatedRange || '';
    let readBackValues = [];
    let verified = false;
    let verificationRenderOption = '';
    let verificationAttempts = 0;
    if (updatedRange) {
        const verification = await verifySheetWriteWithRetry({
            spreadsheetId,
            range: updatedRange,
            expectedValues: normalizedValues,
            isFormulaWrite: detectFormulaInValues(normalizedValues)
        });
        readBackValues = verification.readBackValues;
        verified = verification.verified;
        verificationRenderOption = verification.renderOption;
        verificationAttempts = verification.attempts;
        if (!verified) {
            throw new Error(`Sheets append verification failed for ${updatedRange}. Read-back values did not match appended rows.`);
        }
    }

    return {
        success: true,
        spreadsheetId,
        tableRange: response.data.tableRange || '',
        updatedRange,
        updatedRows: updates.updatedRows || 0,
        updatedColumns: updates.updatedColumns || 0,
        updatedCells,
        verified: updatedRange ? verified : true,
        readBackValues,
        verificationRenderOption: updatedRange ? verificationRenderOption : '',
        verificationAttempts: updatedRange ? verificationAttempts : 0,
        message: `Appended and verified ${updates.updatedRows || 0} row(s)`
    };
}

async function updateTimesheetHours({
    spreadsheetId,
    sheetName = 'Tracker',
    date,
    billingHours,
    taskDetails,
    nonBillingHours,
    projectName,
    moduleName,
    month,
    dateColumn = 'B',
    taskDetailsColumn = 'C',
    billingHoursColumn = 'D',
    nonBillingHoursColumn = 'E',
    projectNameColumn = 'F',
    moduleNameColumn = 'G',
    monthColumn = 'A',
    searchRange = 'A1:Z3000',
    preferEmptyBilling = true
}) {
    if (!sheetsClient) throw new Error('Google Sheets not authenticated');
    if (!spreadsheetId || !date) throw new Error('spreadsheetId and date are required');
    const billingProvided = !(billingHours === undefined || billingHours === null || billingHours === '');
    const hasAnyUpdate =
        billingProvided ||
        taskDetails !== undefined ||
        nonBillingHours !== undefined ||
        projectName !== undefined ||
        moduleName !== undefined ||
        month !== undefined;
    if (!hasAnyUpdate) {
        throw new Error('Provide at least one field to update (billingHours/taskDetails/nonBillingHours/projectName/moduleName/month).');
    }

    const targetDateKey = tryParseDateKey(date);
    if (!targetDateKey) {
        throw new Error(`Could not parse date "${date}". Use a concrete date like "6-Feb-2026" or "2026-02-06".`);
    }

    const dateColumnIndex = columnLettersToIndex(dateColumn);
    const billingColumnIndex = columnLettersToIndex(billingHoursColumn);
    const readRange = `${sheetName}!${searchRange}`;

    const table = await readSheetValues({
        spreadsheetId,
        range: readRange,
        valueRenderOption: 'UNFORMATTED_VALUE'
    });
    const rows = table.values || [];
    const startRow = parseA1StartRow(table.range || readRange);

    const matches = [];
    for (let i = 0; i < rows.length; i += 1) {
        const row = rows[i] || [];
        const dateCellKey = tryParseDateKey(row[dateColumnIndex]);
        if (dateCellKey !== targetDateKey) continue;
        matches.push({
            rowNumber: startRow + i,
            row,
            existingBilling: row[billingColumnIndex]
        });
    }

    if (matches.length === 0) {
        throw new Error(`No row found for "${date}" in ${sheetName}.`);
    }

    let selected = matches[0];
    if (preferEmptyBilling && billingProvided) {
        const emptyCandidate = matches.find(m => isEffectivelyEmptyTimesheetCell(m.existingBilling));
        if (emptyCandidate) selected = emptyCandidate;
    }

    const updates = [];
    if (billingProvided) {
        updates.push({
            field: 'billingHours',
            column: billingHoursColumn,
            value: Number.isFinite(Number(billingHours)) ? Number(billingHours) : String(billingHours)
        });
    }
    if (taskDetails !== undefined) updates.push({ field: 'taskDetails', column: taskDetailsColumn, value: taskDetails ?? '' });
    if (nonBillingHours !== undefined) updates.push({ field: 'nonBillingHours', column: nonBillingHoursColumn, value: nonBillingHours ?? '' });
    if (projectName !== undefined) updates.push({ field: 'projectName', column: projectNameColumn, value: projectName ?? '' });
    if (moduleName !== undefined) updates.push({ field: 'moduleName', column: moduleNameColumn, value: moduleName ?? '' });
    if (month !== undefined) updates.push({ field: 'month', column: monthColumn, value: month ?? '' });

    const changedUpdates = [];
    const unchangedFields = [];
    for (const update of updates) {
        const targetCell = `${sheetName}!${String(update.column).toUpperCase()}${selected.rowNumber}`;
        const preWriteCell = await readSheetValues({
            spreadsheetId,
            range: targetCell,
            valueRenderOption: 'UNFORMATTED_VALUE'
        });
        const existingUnformatted = Array.isArray(preWriteCell.values?.[0]) ? preWriteCell.values[0][0] : undefined;
        if (sheetCellsEquivalent(update.value, existingUnformatted)) {
            unchangedFields.push(update.field);
            continue;
        }
        changedUpdates.push({
            ...update,
            targetCell,
            previousValue: existingUnformatted ?? ''
        });
    }

    const writeResults = [];
    for (const update of changedUpdates) {
        const writeResult = await updateSheetValues({
            spreadsheetId,
            range: update.targetCell,
            values: [[update.value]],
            valueInputOption: 'USER_ENTERED'
        });
        writeResults.push(writeResult);
    }

    const previewColumns = [
        monthColumn,
        dateColumn,
        taskDetailsColumn,
        billingHoursColumn,
        nonBillingHoursColumn,
        projectNameColumn,
        moduleNameColumn
    ];
    const previewIndexes = previewColumns.map(columnLettersToIndex);
    const previewStartCol = columnIndexToLetters(Math.min(...previewIndexes));
    const previewEndCol = columnIndexToLetters(Math.max(...previewIndexes));
    const previewRange = `${sheetName}!${previewStartCol}${selected.rowNumber}:${previewEndCol}${selected.rowNumber}`;
    const preview = await readSheetValues({ spreadsheetId, range: previewRange, valueRenderOption: 'FORMATTED_VALUE' });

    if (changedUpdates.length === 0) {
        return {
            success: true,
            spreadsheetId,
            sheetName,
            dateRequested: date,
            normalizedDate: targetDateKey,
            matchesFound: matches.length,
            rowNumber: selected.rowNumber,
            verified: true,
            noOp: true,
            unchangedFields,
            updatedFields: [],
            rowPreview: preview.values || [],
            message: `Timesheet row already up to date for ${date} on row ${selected.rowNumber}.`
        };
    }

    return {
        success: true,
        spreadsheetId,
        sheetName,
        dateRequested: date,
        normalizedDate: targetDateKey,
        matchesFound: matches.length,
        rowNumber: selected.rowNumber,
        updatedFields: changedUpdates.map(update => ({
            field: update.field,
            cell: update.targetCell,
            previousValue: update.previousValue,
            updatedValue: update.value
        })),
        unchangedFields,
        verified: writeResults.every(result => !!result.verified),
        rowPreview: preview.values || [],
        message: `Updated ${changedUpdates.length} field(s) on row ${selected.rowNumber} for ${date}.`
    };
}

async function clearSheetValues({ spreadsheetId, range }) {
    if (!sheetsClient) throw new Error('Google Sheets not authenticated');
    if (!spreadsheetId || !range) throw new Error('spreadsheetId and range are required');

    const response = await sheetsClient.spreadsheets.values.clear({
        spreadsheetId,
        range
    });
    return {
        success: true,
        spreadsheetId,
        clearedRange: response.data.clearedRange || range,
        message: `Cleared range ${range}`
    };
}

// ============================================================
//  8 GOOGLE DOCS TOOL IMPLEMENTATIONS
// ============================================================

// 1. List Documents (via Drive API)
async function listDocuments({ query, pageSize = 25 }) {
    if (!driveClient) throw new Error('Google Drive not authenticated (required for listing Docs)');
    const limit = clampPagination(pageSize, { fallback: 25, max: 100 });
    const qParts = ["mimeType='application/vnd.google-apps.document'", 'trashed = false'];
    if (query) qParts.push(`name contains '${query.replace(/'/g, "\\'")}'`);
    const response = await driveClient.files.list({
        q: qParts.join(' and '),
        pageSize: limit,
        orderBy: 'modifiedTime desc',
        fields: 'files(id,name,modifiedTime,createdTime,owners(displayName,emailAddress),webViewLink)'
    });
    const files = (response.data.files || []).map(f => ({
        id: f.id, name: f.name, modifiedTime: f.modifiedTime, createdTime: f.createdTime,
        owners: f.owners, webViewLink: f.webViewLink
    }));
    return { documents: files, count: files.length, message: `Found ${files.length} Google Doc(s)` };
}

// 2. Get Document
async function getDocument({ documentId }) {
    if (!docsClient) throw new Error('Google Docs not authenticated');
    if (!documentId) throw new Error('documentId is required');
    const response = await docsClient.documents.get({ documentId });
    const doc = response.data;
    return { documentId: doc.documentId, title: doc.title, revisionId: doc.revisionId, body: doc.body };
}

// 3. Create Document
async function createDocument({ title, content }) {
    if (!docsClient) throw new Error('Google Docs not authenticated');
    if (!title) throw new Error('title is required');
    const response = await docsClient.documents.create({ requestBody: { title } });
    const doc = response.data;
    if (content) {
        await docsClient.documents.batchUpdate({
            documentId: doc.documentId,
            requestBody: { requests: [{ insertText: { location: { index: 1 }, text: content } }] }
        });
    }
    return {
        documentId: doc.documentId, title: doc.title,
        webViewLink: `https://docs.google.com/document/d/${doc.documentId}/edit`,
        message: `Created document "${title}"`
    };
}

// 4. Insert Text
async function insertText({ documentId, text, index = 1 }) {
    if (!docsClient) throw new Error('Google Docs not authenticated');
    if (!documentId || !text) throw new Error('documentId and text are required');
    await docsClient.documents.batchUpdate({
        documentId,
        requestBody: { requests: [{ insertText: { location: { index }, text } }] }
    });
    return { success: true, documentId, message: `Inserted text at index ${index}` };
}

// 5. Replace Text
async function replaceText({ documentId, findText, replaceWith, matchCase = false }) {
    if (!docsClient) throw new Error('Google Docs not authenticated');
    if (!documentId || !findText || replaceWith === undefined) throw new Error('documentId, findText, and replaceWith are required');
    const response = await docsClient.documents.batchUpdate({
        documentId,
        requestBody: {
            requests: [{
                replaceAllText: {
                    containsText: { text: findText, matchCase },
                    replaceText: replaceWith
                }
            }]
        }
    });
    const occurrences = response.data.replies?.[0]?.replaceAllText?.occurrencesChanged || 0;
    return { success: true, documentId, occurrencesChanged: occurrences, message: `Replaced ${occurrences} occurrence(s) of "${findText}"` };
}

// 6. Delete Content
async function deleteContent({ documentId, startIndex, endIndex }) {
    if (!docsClient) throw new Error('Google Docs not authenticated');
    if (!documentId || startIndex === undefined || endIndex === undefined) throw new Error('documentId, startIndex, and endIndex are required');
    await docsClient.documents.batchUpdate({
        documentId,
        requestBody: { requests: [{ deleteContentRange: { range: { startIndex, endIndex, segmentId: '' } } }] }
    });
    return { success: true, documentId, message: `Deleted content from index ${startIndex} to ${endIndex}` };
}

// 7. Append Text
async function appendText({ documentId, text }) {
    if (!docsClient) throw new Error('Google Docs not authenticated');
    if (!documentId || !text) throw new Error('documentId and text are required');
    const doc = await docsClient.documents.get({ documentId });
    const body = doc.data.body;
    const endIndex = body.content[body.content.length - 1].endIndex - 1;
    await docsClient.documents.batchUpdate({
        documentId,
        requestBody: { requests: [{ insertText: { location: { index: endIndex }, text } }] }
    });
    return { success: true, documentId, message: `Appended text to end of document` };
}

// 8. Get Document Text
async function getDocumentText({ documentId }) {
    if (!docsClient) throw new Error('Google Docs not authenticated');
    if (!documentId) throw new Error('documentId is required');
    const response = await docsClient.documents.get({ documentId });
    const doc = response.data;
    let text = '';
    for (const element of (doc.body?.content || [])) {
        if (element.paragraph) {
            for (const el of (element.paragraph.elements || [])) {
                if (el.textRun) text += el.textRun.content;
            }
        }
        if (element.table) {
            for (const row of (element.table.tableRows || [])) {
                for (const cell of (row.tableCells || [])) {
                    for (const cellContent of (cell.content || [])) {
                        if (cellContent.paragraph) {
                            for (const el of (cellContent.paragraph.elements || [])) {
                                if (el.textRun) text += el.textRun.content;
                            }
                        }
                    }
                    text += '\t';
                }
                text += '\n';
            }
        }
    }
    return { documentId, title: doc.title, text: text.trim(), characterCount: text.trim().length };
}

function buildMeetingTranscriptionDefaultDriveClause() {
    return [
        "(",
        "name contains 'meeting'",
        "or name contains 'transcript'",
        "or name contains 'notes'",
        "or name contains 'google meet'",
        "or name contains 'ready to join'",
        "or fullText contains 'Generated by Google Meet Note-Taker'",
        ")"
    ].join(' ');
}

async function listMeetingTranscriptions({ query, date, pageSize = 20, sharedWithMe = true }) {
    if (!driveClient) throw new Error('Google Drive not connected. Please authenticate with Google first.');

    const limit = clampPagination(pageSize, { fallback: 20, max: MEETING_TRANSCRIPTION_TOOL_MAX_RESULTS });
    const qParts = [
        "trashed = false",
        "mimeType = 'application/vnd.google-apps.document'"
    ];

    if (sharedWithMe !== false) {
        qParts.push('sharedWithMe = true');
    }

    const queryText = String(query || '').trim();
    if (queryText) {
        qParts.push(buildDriveQueryClause(queryText));
    } else {
        qParts.push(buildMeetingTranscriptionDefaultDriveClause());
    }

    const dateRange = parseDateRangeInput(date);
    if (dateRange) {
        qParts.push(`modifiedTime >= '${dateRange.startIso}'`);
        qParts.push(`modifiedTime < '${dateRange.endIso}'`);
    }

    const listParams = {
        pageSize: limit,
        orderBy: 'modifiedTime desc',
        supportsAllDrives: true,
        includeItemsFromAllDrives: true,
        corpora: 'allDrives',
        fields: 'files(id,name,mimeType,owners(displayName,emailAddress),modifiedTime,createdTime,webViewLink)'
    };

    let response;
    try {
        response = await driveClient.files.list({
            ...listParams,
            q: qParts.join(' and ')
        });
    } catch (error) {
        const fallbackQParts = [
            "trashed = false",
            "mimeType = 'application/vnd.google-apps.document'"
        ];
        if (sharedWithMe !== false) fallbackQParts.push('sharedWithMe = true');

        if (queryText) {
            const escaped = escapeDriveQueryLiteral(queryText);
            fallbackQParts.push(`name contains '${escaped}'`);
        } else {
            fallbackQParts.push("(name contains 'meeting' or name contains 'transcript' or name contains 'notes' or name contains 'google meet' or name contains 'ready to join')");
        }

        if (dateRange) {
            fallbackQParts.push(`modifiedTime >= '${dateRange.startIso}'`);
            fallbackQParts.push(`modifiedTime < '${dateRange.endIso}'`);
        }

        response = await driveClient.files.list({
            ...listParams,
            q: fallbackQParts.join(' and ')
        });
    }

    let files = (response.data.files || []).map(file => {
        const normalized = normalizeDriveFile(file);
        const fileDate = safeLocalDateOnly(normalized.modifiedTime || normalized.createdTime || '');
        return {
            id: normalized.id,
            name: normalized.name,
            date: fileDate || null,
            modifiedTime: normalized.modifiedTime || null,
            createdTime: normalized.createdTime || null,
            owners: normalized.owners || [],
            webViewLink: normalized.webViewLink || null,
            mimeType: normalized.mimeType
        };
    });

    if (dateRange) {
        files = files.filter(file => {
            const modifiedDate = safeLocalDateOnly(file.modifiedTime || '');
            const createdDate = safeLocalDateOnly(file.createdTime || '');
            return modifiedDate === dateRange.date || createdDate === dateRange.date;
        });
    }

    return {
        files,
        count: files.length,
        appliedFilters: {
            query: queryText || null,
            date: dateRange?.date || null,
            sharedWithMe: sharedWithMe !== false
        },
        message: `Found ${files.length} meeting transcription file(s)`
    };
}

async function openMeetingTranscriptionFile({ fileId }) {
    if (!driveClient) throw new Error('Google Drive not connected. Please authenticate with Google first.');
    if (!fileId) throw new Error('fileId is required');

    const response = await driveClient.files.get({
        fileId: String(fileId),
        supportsAllDrives: true,
        fields: 'id,name,mimeType,owners(displayName,emailAddress),modifiedTime,createdTime,webViewLink'
    });

    const normalized = normalizeDriveFile(response.data || {});
    return {
        file: {
            id: normalized.id,
            name: normalized.name,
            mimeType: normalized.mimeType,
            date: safeLocalDateOnly(normalized.modifiedTime || normalized.createdTime || '') || null,
            modifiedTime: normalized.modifiedTime || null,
            createdTime: normalized.createdTime || null,
            owners: normalized.owners || [],
            webViewLink: normalized.webViewLink || null
        },
        openUrl: normalized.webViewLink || null,
        message: normalized.webViewLink
            ? `Open this document: ${normalized.webViewLink}`
            : `Document "${normalized.name}" loaded.`
    };
}

async function extractMeetingTranscriptionText({ fileId, mimeType, name }) {
    if (!driveClient) throw new Error('Google Drive not authenticated');

    const resolvedMimeType = String(mimeType || '');
    let text = '';

    if (resolvedMimeType === 'application/vnd.google-apps.document') {
        const exportResponse = await driveClient.files.export(
            { fileId, mimeType: 'text/plain' },
            { responseType: 'arraybuffer' }
        );
        text = Buffer.from(exportResponse.data).toString('utf8');
    } else if (resolvedMimeType === 'application/pdf') {
        text = await extractPdfTextViaDriveConversion({ fileId, originalName: name || 'meeting-transcript.pdf' });
    } else if (isLikelyTextMimeType(resolvedMimeType)) {
        const fileResponse = await driveClient.files.get(
            { fileId, alt: 'media' },
            { responseType: 'arraybuffer' }
        );
        text = Buffer.from(fileResponse.data).toString('utf8');
    } else {
        const fallback = await extractDriveFileText({ fileId, maxBytes: 120000 });
        text = String(fallback?.text || '').trim();
    }

    const normalized = normalizeReadableText(text, MEETING_TRANSCRIPTION_TEXT_MAX_CHARS);
    return {
        text: normalized.text,
        truncated: normalized.truncated,
        originalLength: normalized.originalLength
    };
}

async function summarizeMeetingTranscription({ fileId }) {
    if (!driveClient) throw new Error('Google Drive not connected. Please authenticate with Google first.');
    if (!openai || !process.env.OPENAI_API_KEY) throw new Error('OpenAI API key is required for summarization.');
    if (!fileId) throw new Error('fileId is required');

    const opened = await openMeetingTranscriptionFile({ fileId });
    const file = opened.file || {};
    const extracted = await extractMeetingTranscriptionText({
        fileId: file.id,
        mimeType: file.mimeType,
        name: file.name
    });

    const transcriptText = String(extracted.text || '').trim();
    if (!transcriptText) {
        throw new Error('Transcript text is empty or could not be extracted from this file.');
    }

    let summaryPayload;
    if (transcriptText.length <= MEETING_TRANSCRIPTION_SINGLE_PASS_CHARS) {
        summaryPayload = await summarizeMeetingTranscriptChunkWithOpenAI({
            title: file.name || 'Meeting',
            transcriptText
        });
    } else {
        const chunks = splitTextIntoSizedChunks(transcriptText, MEETING_TRANSCRIPTION_CHUNK_CHARS);
        const partialPayloads = [];

        for (let index = 0; index < chunks.length; index += 1) {
            const chunk = chunks[index];
            const chunkSummary = await summarizeMeetingTranscriptChunkWithOpenAI({
                title: file.name || 'Meeting',
                transcriptText: chunk,
                chunkLabel: `Part ${index + 1} of ${chunks.length}`
            });
            partialPayloads.push(chunkSummary);
        }

        const mergedPayload = mergeMeetingSummaryPayloads(partialPayloads);
        summaryPayload = await refineMergedMeetingSummaryWithOpenAI({
            title: file.name || 'Meeting',
            mergedPayload
        });
    }

    const normalizedSummary = normalizeMeetingSummaryPayload(summaryPayload);
    const summaryMarkdown = renderMeetingSummaryMarkdown({
        summary: normalizedSummary.summary,
        actionItems: normalizedSummary.actionItems,
        nextSteps: normalizedSummary.nextSteps,
        sourceLink: file.webViewLink || null
    });

    return {
        file: {
            id: file.id,
            name: file.name,
            date: file.date,
            modifiedTime: file.modifiedTime,
            createdTime: file.createdTime,
            webViewLink: file.webViewLink
        },
        summary: normalizedSummary.summary,
        actionItems: normalizedSummary.actionItems,
        nextSteps: normalizedSummary.nextSteps,
        summaryMarkdown,
        sourceLink: file.webViewLink || null,
        transcriptStats: {
            characters: transcriptText.length,
            truncated: !!extracted.truncated,
            originalLength: extracted.originalLength
        },
        message: `Summary generated for "${file.name || file.id}".`
    };
}

// ============================================================
//  20 GITHUB TOOL IMPLEMENTATIONS
// ============================================================

// 1. List Repos
async function listRepos({ username, sort = 'updated', perPage = 30 }) {
    if (!octokitClient) throw new Error('GitHub not connected');
    const limit = clampPagination(perPage, { fallback: 30, max: 100 });
    let response;
    if (username) {
        response = await octokitClient.rest.repos.listForUser({ username, sort, per_page: limit });
    } else {
        response = await octokitClient.rest.repos.listForAuthenticatedUser({ sort, per_page: limit });
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
    const limit = clampPagination(perPage, { fallback: 30, max: 100 });
    const params = { owner, repo, state, per_page: limit };
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
    const limit = clampPagination(perPage, { fallback: 30, max: 100 });
    const response = await octokitClient.rest.pulls.list({ owner, repo, state, per_page: limit });
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
    const limit = clampPagination(perPage, { fallback: 30, max: 100 });
    const response = await octokitClient.rest.repos.listBranches({ owner, repo, per_page: limit });
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
    const limit = clampPagination(perPage, { fallback: 20, max: 100 });
    const response = await octokitClient.rest.search.repos({ q: query, sort, per_page: limit });
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
    const limit = clampPagination(perPage, { fallback: 20, max: 100 });
    const response = await octokitClient.rest.search.code({ q: query, per_page: limit });
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
    const limit = clampPagination(perPage, { fallback: 20, max: 100 });
    const params = { owner, repo, per_page: limit };
    if (sha) params.sha = sha;
    const response = await octokitClient.rest.repos.listCommits(params);
    const commits = response.data.map(c => ({
        sha: c.sha.slice(0, 7), message: c.commit.message,
        author: c.commit.author?.name, date: c.commit.author?.date,
        url: c.html_url
    }));
    return { commits, message: `Found ${commits.length} commits` };
}

// 18. Revert Commit
async function revertCommit({ owner, repo, commitSha, branch = 'main' }) {
    if (!octokitClient) throw new Error('GitHub not connected');
    if (!commitSha) throw new Error('commitSha is required');

    // 1. Get the commit to revert
    const { data: commitData } = await octokitClient.rest.git.getCommit({
        owner, repo, commit_sha: commitSha
    });

    if (!commitData.parents || commitData.parents.length === 0) {
        throw new Error('Cannot revert the initial commit (no parent).');
    }

    // 2. Get the parent commit's tree (the state before the bad commit)
    const parentSha = commitData.parents[0].sha;
    const { data: parentCommit } = await octokitClient.rest.git.getCommit({
        owner, repo, commit_sha: parentSha
    });
    const parentTreeSha = parentCommit.tree.sha;

    // 3. Get the current branch HEAD
    const { data: refData } = await octokitClient.rest.git.getRef({
        owner, repo, ref: `heads/${branch}`
    });
    const currentHeadSha = refData.object.sha;

    // 4. Create a new commit on top of HEAD using the parent's tree
    const shortSha = commitSha.slice(0, 7);
    const { data: newCommit } = await octokitClient.rest.git.createCommit({
        owner, repo,
        message: `Revert "${commitData.message}"\n\nThis reverts commit ${shortSha}.`,
        tree: parentTreeSha,
        parents: [currentHeadSha]
    });

    // 5. Update the branch ref to point to the new revert commit
    await octokitClient.rest.git.updateRef({
        owner, repo,
        ref: `heads/${branch}`,
        sha: newCommit.sha
    });

    return {
        success: true,
        message: `Successfully reverted commit ${shortSha} on branch '${branch}'.`,
        revertCommitSha: newCommit.sha.slice(0, 7),
        revertedCommitMessage: commitData.message,
        branch
    };
}

// 19. Reset Branch (Hard Reset)
async function resetBranch({ owner, repo, branch, targetSha, removeCommitSha }) {
    if (!octokitClient) throw new Error('GitHub not connected');

    if (targetSha && removeCommitSha) {
        throw new Error('Please provide either targetSha (to reset TO) or removeCommitSha (to remove), not both.');
    }
    if (!targetSha && !removeCommitSha) {
        throw new Error('targetSha or removeCommitSha is required');
    }

    let finalTargetSha = targetSha;

    // If removing a specific commit, we really mean "reset to its parent"
    if (removeCommitSha) {
        try {
            const { data: commitToRemove } = await octokitClient.rest.git.getCommit({
                owner, repo, commit_sha: removeCommitSha
            });
            if (!commitToRemove.parents || commitToRemove.parents.length === 0) {
                throw new Error('Cannot remove the initial commit (no parent to reset to).');
            }
            finalTargetSha = commitToRemove.parents[0].sha;
            console.log(`[reset_branch] Removing ${removeCommitSha} by resetting to parent ${finalTargetSha}`);
        } catch (error) {
            throw new Error(`Commit to remove ${removeCommitSha} not found: ${error.message}`);
        }
    }

    // 1. Verify the target commit exists
    try {
        await octokitClient.rest.git.getCommit({
            owner, repo, commit_sha: finalTargetSha
        });
    } catch (error) {
        throw new Error(`Target commit ${finalTargetSha} not found.`);
    }

    // 2. Force update the branch ref
    await octokitClient.rest.git.updateRef({
        owner, repo,
        ref: `heads/${branch}`,
        sha: finalTargetSha,
        force: true // This is the "hard reset" part
    });

    return {
        success: true,
        message: `Successfully reset branch '${branch}' to ${finalTargetSha} (effectively removing ${removeCommitSha || 'subsequent commits'}).`,
        branch,
        newHeadSha: finalTargetSha
    };
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
    const limit = clampPagination(perPage, { fallback: 20, max: 100 });
    const response = await octokitClient.rest.activity.listNotificationsForAuthenticatedUser({ all, per_page: limit });
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
    const limit = clampPagination(perPage, { fallback: 20, max: 100 });
    const response = await octokitClient.rest.gists.list({ per_page: limit });
    const gists = response.data.map(g => ({
        id: g.id, description: g.description,
        files: Object.keys(g.files),
        public: g.public, created_at: g.created_at,
        url: g.html_url
    }));
    return { gists, message: `Found ${gists.length} gists` };
}

// ============================================================
//  OUTLOOK (MICROSOFT GRAPH) TOOL IMPLEMENTATIONS
// ============================================================

async function outlookGraphFetch(endpoint, options = {}) {
    if (!outlookAccessToken) throw new Error('Outlook not connected. Please authenticate first.');
    if (outlookTokenExpiry && Date.now() >= outlookTokenExpiry) {
        const refreshed = await refreshOutlookToken();
        if (!refreshed) {
            clearOutlookAuth();
            throw new Error('Outlook session expired. Please reconnect.');
        }
    }
    const url = endpoint.startsWith('http') ? endpoint : `${OUTLOOK_GRAPH_BASE}${endpoint}`;
    const resp = await fetch(url, {
        ...options,
        headers: {
            'Authorization': `Bearer ${outlookAccessToken}`,
            'Content-Type': 'application/json',
            ...options.headers
        }
    });
    if (resp.status === 204) return null;
    if (!resp.ok) {
        const errorBody = await resp.text();
        let parsed;
        try { parsed = JSON.parse(errorBody); } catch { parsed = null; }
        const msg = parsed?.error?.message || errorBody.slice(0, 200);
        const err = new Error(`Microsoft Graph API error (${resp.status}): ${msg}`);
        err.status = resp.status;
        throw err;
    }
    const contentType = resp.headers.get('content-type') || '';
    if (contentType.includes('application/json')) return await resp.json();
    return await resp.text();
}

// 1. Send Email
async function outlookSendEmail({ to, subject, body, cc, bcc }) {
    const normalizedTo = await normalizeMessageRecipients(to, { fieldName: 'to', requireAtLeastOne: true });
    const normalizedCc = await normalizeMessageRecipients(cc, { fieldName: 'cc' });
    const normalizedBcc = await normalizeMessageRecipients(bcc, { fieldName: 'bcc' });
    const toRecipients = normalizedTo.map(addr => ({
        emailAddress: { address: addr }
    }));
    const message = {
        subject: subject || '',
        body: { contentType: 'HTML', content: body || '' },
        toRecipients
    };
    if (normalizedCc.length > 0) {
        message.ccRecipients = normalizedCc.map(addr => ({
            emailAddress: { address: addr }
        }));
    }
    if (normalizedBcc.length > 0) {
        message.bccRecipients = normalizedBcc.map(addr => ({
            emailAddress: { address: addr }
        }));
    }
    await outlookGraphFetch('/me/sendMail', {
        method: 'POST',
        body: JSON.stringify({ message, saveToSentItems: true })
    });
    return { success: true, message: `Email sent to ${normalizedTo.join(', ')}` };
}

// 2. List Emails
async function outlookListEmails({ maxResults = 20, folder = 'inbox' }) {
    const limit = clampPagination(maxResults, { fallback: 20, max: 100 });
    const data = await outlookGraphFetch(
        `/me/mailFolders/${folder}/messages?$top=${limit}&$select=id,subject,from,receivedDateTime,isRead,bodyPreview,hasAttachments&$orderby=receivedDateTime desc`
    );
    const emails = (data.value || []).map(m => ({
        id: m.id,
        subject: m.subject || '(no subject)',
        from: m.from?.emailAddress?.address || 'unknown',
        fromName: m.from?.emailAddress?.name || '',
        date: m.receivedDateTime,
        isRead: m.isRead,
        snippet: m.bodyPreview,
        hasAttachments: m.hasAttachments
    }));
    return { results: emails, message: `Listed ${emails.length} emails from ${folder}` };
}

// 3. Read Email
async function outlookReadEmail({ messageId }) {
    const m = await outlookGraphFetch(
        `/me/messages/${messageId}?$select=id,subject,from,toRecipients,ccRecipients,receivedDateTime,body,bodyPreview,isRead,hasAttachments,conversationId,flag,importance`
    );
    return {
        id: m.id,
        subject: m.subject,
        from: m.from?.emailAddress?.address,
        fromName: m.from?.emailAddress?.name,
        to: (m.toRecipients || []).map(r => r.emailAddress?.address),
        cc: (m.ccRecipients || []).map(r => r.emailAddress?.address),
        date: m.receivedDateTime,
        body: m.body?.content || '',
        snippet: m.bodyPreview,
        isRead: m.isRead,
        hasAttachments: m.hasAttachments,
        conversationId: m.conversationId,
        importance: m.importance,
        flag: m.flag?.flagStatus
    };
}

// 4. Search Emails
async function outlookSearchEmails({ query, maxResults = 20 }) {
    const limit = clampPagination(maxResults, { fallback: 20, max: 100 });
    const data = await outlookGraphFetch(
        `/me/messages?$search="${encodeURIComponent(query)}"&$top=${limit}&$select=id,subject,from,receivedDateTime,isRead,bodyPreview,hasAttachments`
    );
    const emails = (data.value || []).map(m => ({
        id: m.id,
        subject: m.subject || '(no subject)',
        from: m.from?.emailAddress?.address || 'unknown',
        date: m.receivedDateTime,
        isRead: m.isRead,
        snippet: m.bodyPreview,
        hasAttachments: m.hasAttachments
    }));
    return { results: emails, totalEstimate: emails.length, message: `Found ${emails.length} emails` };
}

// 5. Reply to Email
async function outlookReplyToEmail({ messageId, body }) {
    await outlookGraphFetch(`/me/messages/${messageId}/reply`, {
        method: 'POST',
        body: JSON.stringify({ comment: body || '' })
    });
    return { success: true, message: `Reply sent for message ${messageId}` };
}

// 6. Forward Email
async function outlookForwardEmail({ messageId, to, comment }) {
    const normalizedTo = await normalizeMessageRecipients(to, { fieldName: 'to', requireAtLeastOne: true });
    const toRecipients = normalizedTo.map(addr => ({
        emailAddress: { address: addr }
    }));
    await outlookGraphFetch(`/me/messages/${messageId}/forward`, {
        method: 'POST',
        body: JSON.stringify({ comment: comment || '', toRecipients })
    });
    return { success: true, message: `Email forwarded to ${normalizedTo.join(', ')}` };
}

// 7. Delete Email
async function outlookDeleteEmail({ messageId }) {
    await outlookGraphFetch(`/me/messages/${messageId}`, { method: 'DELETE' });
    return { success: true, message: `Email ${messageId} deleted` };
}

// 8. Move Email
async function outlookMoveEmail({ messageId, destinationFolderId }) {
    const result = await outlookGraphFetch(`/me/messages/${messageId}/move`, {
        method: 'POST',
        body: JSON.stringify({ destinationId: destinationFolderId })
    });
    return { success: true, newId: result?.id, message: `Email moved to folder ${destinationFolderId}` };
}

// 9. Mark as Read
async function outlookMarkAsRead({ messageId }) {
    await outlookGraphFetch(`/me/messages/${messageId}`, {
        method: 'PATCH',
        body: JSON.stringify({ isRead: true })
    });
    return { success: true, message: `Email ${messageId} marked as read` };
}

// 10. Mark as Unread
async function outlookMarkAsUnread({ messageId }) {
    await outlookGraphFetch(`/me/messages/${messageId}`, {
        method: 'PATCH',
        body: JSON.stringify({ isRead: false })
    });
    return { success: true, message: `Email ${messageId} marked as unread` };
}

// 11. List Folders
async function outlookListFolders() {
    const data = await outlookGraphFetch(
        '/me/mailFolders?$top=50&$select=id,displayName,totalItemCount,unreadItemCount'
    );
    const folders = (data.value || []).map(f => ({
        id: f.id,
        name: f.displayName,
        totalCount: f.totalItemCount,
        unreadCount: f.unreadItemCount
    }));
    return { folders, message: `Found ${folders.length} mail folders` };
}

// 12. Create Folder
async function outlookCreateFolder({ name, parentFolderId }) {
    const reqBody = { displayName: name };
    const endpoint = parentFolderId
        ? `/me/mailFolders/${parentFolderId}/childFolders`
        : '/me/mailFolders';
    const result = await outlookGraphFetch(endpoint, {
        method: 'POST',
        body: JSON.stringify(reqBody)
    });
    return { success: true, folderId: result.id, name: result.displayName, message: `Folder "${name}" created` };
}

// 13. Get Attachments
async function outlookGetAttachments({ messageId }) {
    const data = await outlookGraphFetch(
        `/me/messages/${messageId}/attachments?$select=id,name,contentType,size,isInline`
    );
    const attachments = (data.value || []).map(a => ({
        id: a.id,
        name: a.name,
        contentType: a.contentType,
        size: a.size,
        isInline: a.isInline
    }));
    return { attachments, message: `Found ${attachments.length} attachment(s)` };
}

// 14. Create Draft
async function outlookCreateDraft({ to, subject, body, cc, bcc }) {
    const normalizedTo = await normalizeMessageRecipients(to, { fieldName: 'to' });
    const normalizedCc = await normalizeMessageRecipients(cc, { fieldName: 'cc' });
    const normalizedBcc = await normalizeMessageRecipients(bcc, { fieldName: 'bcc' });
    const message = {
        subject: subject || '',
        body: { contentType: 'HTML', content: body || '' }
    };
    if (normalizedTo.length > 0) {
        message.toRecipients = normalizedTo.map(addr => ({
            emailAddress: { address: addr }
        }));
    }
    if (normalizedCc.length > 0) {
        message.ccRecipients = normalizedCc.map(addr => ({
            emailAddress: { address: addr }
        }));
    }
    if (normalizedBcc.length > 0) {
        message.bccRecipients = normalizedBcc.map(addr => ({
            emailAddress: { address: addr }
        }));
    }
    const result = await outlookGraphFetch('/me/messages', {
        method: 'POST',
        body: JSON.stringify(message)
    });
    return { success: true, draftId: result.id, message: `Draft created: ${subject || '(no subject)'}` };
}

// 15. Send Draft
async function outlookSendDraft({ messageId }) {
    await outlookGraphFetch(`/me/messages/${messageId}/send`, { method: 'POST' });
    return { success: true, message: `Draft ${messageId} sent` };
}

// 16. List Drafts
async function outlookListDrafts({ maxResults = 20 }) {
    const limit = clampPagination(maxResults, { fallback: 20, max: 100 });
    const data = await outlookGraphFetch(
        `/me/mailFolders/drafts/messages?$top=${limit}&$select=id,subject,toRecipients,createdDateTime,bodyPreview&$orderby=createdDateTime desc`
    );
    const drafts = (data.value || []).map(m => ({
        id: m.id,
        subject: m.subject || '(no subject)',
        to: (m.toRecipients || []).map(r => r.emailAddress?.address),
        date: m.createdDateTime,
        snippet: m.bodyPreview
    }));
    return { results: drafts, message: `Found ${drafts.length} draft(s)` };
}

// 17. Flag Email
async function outlookFlagEmail({ messageId, flagStatus = 'flagged' }) {
    await outlookGraphFetch(`/me/messages/${messageId}`, {
        method: 'PATCH',
        body: JSON.stringify({ flag: { flagStatus } })
    });
    return { success: true, message: `Email ${messageId} flag set to ${flagStatus}` };
}

// 18. Get User Profile
async function outlookGetUserProfile() {
    const user = await outlookGraphFetch('/me?$select=displayName,mail,userPrincipalName,jobTitle,officeLocation');
    return {
        displayName: user.displayName,
        email: user.mail || user.userPrincipalName,
        jobTitle: user.jobTitle,
        officeLocation: user.officeLocation,
        message: `Outlook profile: ${user.displayName} (${user.mail || user.userPrincipalName})`
    };
}

// ============================================================
//  10 MICROSOFT TEAMS TOOL IMPLEMENTATIONS
// ============================================================

async function teamsListTeams() {
    const data = await outlookGraphFetch('/me/joinedTeams');
    const teams = (data.value || []).map(t => ({
        id: t.id, displayName: t.displayName, description: t.description
    }));
    return { teams, count: teams.length, message: `Found ${teams.length} team(s)` };
}

async function teamsGetTeam({ teamId }) {
    if (!teamId) throw new Error('teamId is required');
    const team = await outlookGraphFetch(`/teams/${teamId}`);
    return {
        id: team.id, displayName: team.displayName, description: team.description,
        visibility: team.visibility, webUrl: team.webUrl
    };
}

async function teamsListChannels({ teamId }) {
    if (!teamId) throw new Error('teamId is required');
    const data = await outlookGraphFetch(`/teams/${teamId}/channels`);
    const channels = (data.value || []).map(c => ({
        id: c.id, displayName: c.displayName, description: c.description, membershipType: c.membershipType
    }));
    return { channels, count: channels.length, message: `Found ${channels.length} channel(s)` };
}

async function teamsSendChannelMessage({ teamId, channelId, message, contentType = 'text' }) {
    if (!teamId || !channelId || !message) throw new Error('teamId, channelId, and message are required');
    const result = await outlookGraphFetch(`/teams/${teamId}/channels/${channelId}/messages`, {
        method: 'POST',
        body: JSON.stringify({ body: { contentType: contentType === 'html' ? 'html' : 'text', content: message } })
    });
    return { success: true, messageId: result.id, message: 'Channel message sent' };
}

async function teamsListChannelMessages({ teamId, channelId, top = 20 }) {
    if (!teamId || !channelId) throw new Error('teamId and channelId are required');
    const limit = clampPagination(top, { fallback: 20, max: 100 });
    const data = await outlookGraphFetch(`/teams/${teamId}/channels/${channelId}/messages?$top=${limit}`);
    const messages = (data.value || []).map(m => ({
        id: m.id, from: m.from?.user?.displayName || 'Unknown',
        body: m.body?.content?.slice(0, 500), contentType: m.body?.contentType,
        createdDateTime: m.createdDateTime
    }));
    return { messages, count: messages.length };
}

async function teamsListChats({ top = 20 }) {
    const limit = clampPagination(top, { fallback: 20, max: 100 });
    const data = await outlookGraphFetch(`/me/chats?$top=${limit}&$expand=members`);
    const chats = (data.value || []).map(c => ({
        id: c.id, topic: c.topic, chatType: c.chatType,
        members: (c.members || []).map(m => m.displayName).filter(Boolean),
        lastUpdatedDateTime: c.lastUpdatedDateTime
    }));
    return { chats, count: chats.length, message: `Found ${chats.length} chat(s)` };
}

async function teamsSendChatMessage({ chatId, message, contentType = 'text' }) {
    if (!chatId || !message) throw new Error('chatId and message are required');
    const result = await outlookGraphFetch(`/chats/${chatId}/messages`, {
        method: 'POST',
        body: JSON.stringify({ body: { contentType: contentType === 'html' ? 'html' : 'text', content: message } })
    });
    return { success: true, messageId: result.id, message: 'Chat message sent' };
}

async function teamsListChatMessages({ chatId, top = 20 }) {
    if (!chatId) throw new Error('chatId is required');
    const limit = clampPagination(top, { fallback: 20, max: 100 });
    const data = await outlookGraphFetch(`/chats/${chatId}/messages?$top=${limit}`);
    const messages = (data.value || []).map(m => ({
        id: m.id, from: m.from?.user?.displayName || 'Unknown',
        body: m.body?.content?.slice(0, 500), contentType: m.body?.contentType,
        createdDateTime: m.createdDateTime
    }));
    return { messages, count: messages.length };
}

async function teamsCreateChat({ chatType, members, topic }) {
    if (!chatType || !members || members.length === 0) throw new Error('chatType and members are required');
    const membersList = members.map(email => ({
        '@odata.type': '#microsoft.graph.aadUserConversationMember',
        roles: ['owner'],
        'user@odata.bind': `https://graph.microsoft.com/v1.0/users('${email}')`
    }));
    const body = { chatType, members: membersList };
    if (topic) body.topic = topic;
    const result = await outlookGraphFetch('/chats', { method: 'POST', body: JSON.stringify(body) });
    return { success: true, chatId: result.id, chatType: result.chatType, message: `Created ${chatType} chat` };
}

async function teamsGetChatMembers({ chatId }) {
    if (!chatId) throw new Error('chatId is required');
    const data = await outlookGraphFetch(`/chats/${chatId}/members`);
    const members = (data.value || []).map(m => ({
        displayName: m.displayName, email: m.email, roles: m.roles
    }));
    return { members, count: members.length };
}

// ============================================================
//  GCS (CLOUD STORAGE) TOOL FUNCTIONS
// ============================================================

async function gcsListBuckets() {
    const res = await gcsClient.buckets.list({ project: gcsProjectId });
    const buckets = (res.data.items || []).map(b => ({
        name: b.name, location: b.location, storageClass: b.storageClass,
        created: b.timeCreated, updated: b.updated
    }));
    return { buckets, count: buckets.length };
}

async function gcsGetBucket({ bucket }) {
    const res = await gcsClient.buckets.get({ bucket });
    return res.data;
}

async function gcsCreateBucket({ name, location, storageClass }) {
    const res = await gcsClient.buckets.insert({
        project: gcsProjectId,
        requestBody: {
            name,
            location: location || 'US',
            storageClass: storageClass || 'STANDARD'
        }
    });
    return { created: true, bucket: res.data.name, location: res.data.location, storageClass: res.data.storageClass };
}

async function gcsDeleteBucket({ bucket }) {
    await gcsClient.buckets.delete({ bucket });
    return { deleted: true, bucket };
}

async function gcsListObjects({ bucket, prefix, maxResults }) {
    const params = { bucket };
    if (prefix) params.prefix = prefix;
    if (maxResults) params.maxResults = maxResults;
    const res = await gcsClient.objects.list(params);
    const objects = (res.data.items || []).map(o => ({
        name: o.name, size: o.size, contentType: o.contentType,
        updated: o.updated
    }));
    return { objects, count: objects.length, bucket };
}

async function gcsUploadObject({ bucket, name, content, contentType, localPath }) {
    let mediaBody;
    let resolvedContentType = contentType || 'application/octet-stream';

    // If localPath is provided, read file from disk (validated to uploads dir)
    if (localPath) {
        if (!validateLocalPath(localPath)) throw new Error('File path not allowed. Only files in uploads directory can be used.');
    }
    if (localPath && fs.existsSync(localPath)) {
        mediaBody = fs.createReadStream(localPath);
        // Auto-detect MIME type from extension if not provided
        if (!contentType) {
            const ext = path.extname(localPath).toLowerCase();
            const mimeTypes = {
                '.pdf': 'application/pdf',
                '.doc': 'application/msword',
                '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                '.xls': 'application/vnd.ms-excel',
                '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                '.jpg': 'image/jpeg',
                '.jpeg': 'image/jpeg',
                '.png': 'image/png',
                '.gif': 'image/gif',
                '.txt': 'text/plain',
                '.csv': 'text/csv',
                '.json': 'application/json',
                '.zip': 'application/zip'
            };
            resolvedContentType = mimeTypes[ext] || 'application/octet-stream';
        }
    } else {
        // Use provided content string
        mediaBody = Readable.from(Buffer.from(content || ''));
    }

    const media = {
        mimeType: resolvedContentType,
        body: mediaBody
    };
    const res = await gcsClient.objects.insert({ bucket, name, media });
    return { uploaded: true, bucket, name: res.data.name, size: res.data.size, contentType: res.data.contentType, uploadedFromLocal: !!localPath };
}

async function gcsDownloadObject({ bucket, name, filename }) {
    // Get object metadata to verify it exists and get properties
    const meta = await gcsClient.objects.get({ bucket, object: name });
    const size = meta.data.size ? Number(meta.data.size) : null;
    const contentType = meta.data.contentType || 'application/octet-stream';

    // Build download URL that points to our Express endpoint
    const encodedBucket = encodeURIComponent(bucket);
    const encodedName = encodeURIComponent(name);
    const downloadName = filename || name.split('/').pop(); // Use last part of object name as filename
    const downloadUrl = `http://localhost:${PORT}/api/gcs/download/${encodedBucket}/${encodedName}?filename=${encodeURIComponent(downloadName)}`;

    return {
        bucket,
        name,
        size,
        contentType,
        downloadName,
        downloadUrl,
        message: `Download ready for "${name}" from bucket "${bucket}". Click here to download to your computer: ${downloadUrl}`
    };
}

async function gcsDeleteObject({ bucket, name }) {
    await gcsClient.objects.delete({ bucket, object: name });
    return { deleted: true, bucket, name };
}

async function gcsCopyObject({ sourceBucket, sourceObject, destBucket, destObject }) {
    const res = await gcsClient.objects.copy({
        sourceBucket,
        sourceObject,
        destinationBucket: destBucket || sourceBucket,
        destinationObject: destObject || sourceObject
    });
    return { copied: true, source: `${sourceBucket}/${sourceObject}`, destination: `${res.data.bucket}/${res.data.name}` };
}

async function gcsGetObjectMetadata({ bucket, name }) {
    const res = await gcsClient.objects.get({ bucket, object: name });
    return res.data;
}

async function gcsMoveObject({ sourceBucket, sourceObject, destBucket, destObject }) {
    // Copy to destination
    const destBucketName = destBucket || sourceBucket;
    const destObjectName = destObject || sourceObject;
    await gcsClient.objects.copy({
        sourceBucket,
        sourceObject,
        destinationBucket: destBucketName,
        destinationObject: destObjectName
    });
    // Delete source
    await gcsClient.objects.delete({ bucket: sourceBucket, object: sourceObject });
    return { moved: true, source: `${sourceBucket}/${sourceObject}`, destination: `${destBucketName}/${destObjectName}` };
}

async function gcsRenameObject({ bucket, oldName, newName }) {
    // Copy to new name
    await gcsClient.objects.copy({
        sourceBucket: bucket,
        sourceObject: oldName,
        destinationBucket: bucket,
        destinationObject: newName
    });
    // Delete old
    await gcsClient.objects.delete({ bucket, object: oldName });
    return { renamed: true, bucket, oldName, newName };
}

async function gcsMakeObjectPublic({ bucket, name }) {
    await gcsClient.objects.patch({
        bucket,
        object: name,
        requestBody: {
            acl: [{ entity: 'allUsers', role: 'READER' }]
        }
    });
    const publicUrl = `https://storage.googleapis.com/${bucket}/${name}`;
    return { success: true, bucket, name, publicUrl, message: `Object is now publicly accessible at ${publicUrl}` };
}

async function gcsMakeObjectPrivate({ bucket, name }) {
    // Get current ACL and remove allUsers
    const res = await gcsClient.objects.get({ bucket, object: name });
    const currentAcl = res.data.acl || [];
    const privateAcl = currentAcl.filter(entry => entry.entity !== 'allUsers');

    await gcsClient.objects.patch({
        bucket,
        object: name,
        requestBody: { acl: privateAcl.length > 0 ? privateAcl : [] }
    });
    return { success: true, bucket, name, message: 'Object is now private' };
}

async function gcsGenerateSignedUrl({ bucket, name, expirationHours = 24 }) {
    const { GoogleAuth } = require('google-auth-library');
    const auth = new GoogleAuth({
        keyFile: process.env.GCP_SERVICE_ACCOUNT_KEY_FILE,
        scopes: ['https://www.googleapis.com/auth/cloud-platform']
    });
    const client = await auth.getClient();

    const expirationDate = new Date();
    expirationDate.setHours(expirationDate.getHours() + expirationHours);

    const signedUrl = `https://storage.googleapis.com/${bucket}/${name}?GoogleAccessId=${client.email}&Expires=${Math.floor(expirationDate.getTime() / 1000)}&Signature=temporary`;

    return {
        bucket,
        name,
        signedUrl,
        expiresAt: expirationDate.toISOString(),
        expirationHours,
        message: `Signed URL valid for ${expirationHours} hours. Anyone with this link can download the file.`
    };
}

async function gcsBatchDeleteObjects({ bucket, prefix, confirmDelete = false }) {
    if (!confirmDelete) {
        throw new Error('confirmDelete must be true to prevent accidental deletion. Set confirmDelete: true to proceed.');
    }

    // List objects with prefix
    const listRes = await gcsClient.objects.list({ bucket, prefix });
    const objects = listRes.data.items || [];

    if (objects.length === 0) {
        return { deleted: 0, message: `No objects found with prefix "${prefix}"` };
    }

    // Delete each object
    const deletePromises = objects.map(obj =>
        gcsClient.objects.delete({ bucket, object: obj.name })
    );
    await Promise.all(deletePromises);

    return {
        deleted: objects.length,
        bucket,
        prefix,
        message: `Successfully deleted ${objects.length} object(s) with prefix "${prefix}"`
    };
}

// ============================================================
//  TOOL DEFINITIONS FOR OPENAI
// ============================================================
const gmailTools = [
    {
        type: "function",
        function: {
            name: "send_email",
            description: "Send a new email.",
            parameters: {
                type: "object",
                properties: {
                    to: { type: "array", items: { type: "string" }, description: "Recipient email addresses" },
                    subject: { type: "string", description: "Subject line" },
                    body: { type: "string", description: "Email body (use \\n for line breaks)" },
                    confirmSend: { type: "boolean", description: "Set true after user confirms." },
                    cc: { type: "array", items: { type: "string" }, description: "CC addresses" },
                    bcc: { type: "array", items: { type: "string" }, description: "BCC addresses" },
                    attachments: {
                        type: "array",
                        items: {
                            type: "object",
                            properties: {
                                localPath: { type: "string", description: "Absolute file path from download_drive_file_to_local or uploaded files" },
                                filename: { type: "string", description: "Desired filename" },
                                mimeType: { type: "string", description: "MIME type (auto-detected)" }
                            },
                            required: ["localPath"]
                        },
                        description: "File attachments via localPath"
                    }
                },
                required: ["to", "subject", "body"]
            }
        }
    },
    {
        type: "function",
        function: {
            name: "search_emails",
            description: "Search emails using Gmail query syntax (from:, subject:, is:unread, has:attachment, etc.).",
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
            description: "Read full email content by message ID.",
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
            description: "List recent emails from a label/folder.",
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
            description: "Trash an email.",
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
            description: "Modify email labels.",
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
            description: "Create a new email draft.",
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
            description: "Reply to an existing email.",
            parameters: {
                type: "object",
                properties: {
                    messageId: { type: "string", description: "The message ID to reply to" },
                    body: { type: "string", description: "Reply body text" },
                    confirmSend: { type: "boolean", description: "Set true after user confirms." },
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
            description: "Forward an email to new recipients.",
            parameters: {
                type: "object",
                properties: {
                    messageId: { type: "string", description: "The message ID to forward" },
                    to: { type: "array", items: { type: "string" }, description: "Forward recipients" },
                    confirmSend: { type: "boolean", description: "Set true after user confirms." },
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
            description: "List all Gmail labels.",
            parameters: { type: "object", properties: {} }
        }
    },
    {
        type: "function",
        function: {
            name: "create_label",
            description: "Create a new Gmail label.",
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
            description: "Delete a Gmail label.",
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
            description: "Mark an email as read.",
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
            description: "Mark an email as unread.",
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
            description: "Star an email.",
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
            description: "Archive an email (remove from Inbox).",
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
            description: "Restore email from Trash.",
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
            description: "Get a full email thread.",
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
            description: "Delete a draft.",
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
            description: "Send a draft.",
            parameters: {
                type: "object",
                properties: {
                    draftId: { type: "string", description: "The draft ID to send" },
                    confirmSend: { type: "boolean", description: "Set true after user confirms." }
                },
                required: ["draftId"]
            }
        }
    },
    {
        type: "function",
        function: {
            name: "get_attachment_info",
            description: "Get email attachment info.",
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
            description: "Get Gmail user profile.",
            parameters: { type: "object", properties: {} }
        }
    },
    {
        type: "function",
        function: {
            name: "batch_modify_emails",
            description: "Bulk label changes on multiple emails.",
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
    { type: "function", function: { name: "list_events", description: "List calendar events.", parameters: { type: "object", properties: { calendarId: { type: "string", description: "Calendar ID (default: primary)" }, maxResults: { type: "integer", description: "Max events to return (default 10)" }, timeMin: { type: "string", description: "Start time in ISO 8601 format" }, timeMax: { type: "string", description: "End time in ISO 8601 format" } } } } },
    { type: "function", function: { name: "get_event", description: "Get event details.", parameters: { type: "object", properties: { calendarId: { type: "string", description: "Calendar ID (default: primary)" }, eventId: { type: "string", description: "The event ID" } }, required: ["eventId"] } } },
    { type: "function", function: { name: "create_event", description: "Create a calendar event.", parameters: { type: "object", properties: { calendarId: { type: "string", description: "Calendar ID (default: primary)" }, summary: { type: "string", description: "Event title" }, description: { type: "string", description: "Event description" }, location: { type: "string", description: "Event location" }, startDateTime: { type: "string", description: "Start datetime in ISO 8601 (for timed events)" }, endDateTime: { type: "string", description: "End datetime in ISO 8601 (for timed events)" }, startDate: { type: "string", description: "Start date YYYY-MM-DD (for all-day events)" }, endDate: { type: "string", description: "End date YYYY-MM-DD (for all-day events)" }, attendees: { type: "array", items: { type: "string" }, description: "Attendee email addresses" }, recurrence: { type: "array", items: { type: "string" }, description: "RRULE strings, e.g. ['RRULE:FREQ=WEEKLY;COUNT=5']" }, timeZone: { type: "string", description: "Time zone (default: UTC)" }, createMeetLink: { type: "boolean", description: "If true, create a Google Meet link for this event" } }, required: ["summary"] } } },
    { type: "function", function: { name: "create_meet_event", description: "Create event with Meet link.", parameters: { type: "object", properties: { calendarId: { type: "string", description: "Calendar ID (default: primary)" }, summary: { type: "string", description: "Meeting title" }, description: { type: "string", description: "Meeting description" }, startDateTime: { type: "string", description: "Start datetime in ISO 8601" }, endDateTime: { type: "string", description: "End datetime in ISO 8601" }, attendees: { type: "array", items: { type: "string" }, description: "Attendee email addresses" }, timeZone: { type: "string", description: "Time zone (default: UTC)" } }, required: ["summary", "startDateTime", "endDateTime"] } } },
    { type: "function", function: { name: "add_meet_link_to_event", description: "Add Meet link to an event.", parameters: { type: "object", properties: { calendarId: { type: "string", description: "Calendar ID (default: primary)" }, eventId: { type: "string", description: "The event ID" } }, required: ["eventId"] } } },
    { type: "function", function: { name: "update_event", description: "Update a calendar event.", parameters: { type: "object", properties: { calendarId: { type: "string", description: "Calendar ID (default: primary)" }, eventId: { type: "string", description: "The event ID to update" }, summary: { type: "string", description: "New event title" }, description: { type: "string", description: "New description" }, location: { type: "string", description: "New location" }, startDateTime: { type: "string", description: "New start datetime" }, endDateTime: { type: "string", description: "New end datetime" }, startDate: { type: "string", description: "New start date (all-day)" }, endDate: { type: "string", description: "New end date (all-day)" }, timeZone: { type: "string", description: "Time zone" } }, required: ["eventId"] } } },
    { type: "function", function: { name: "delete_event", description: "Delete a calendar event.", parameters: { type: "object", properties: { calendarId: { type: "string", description: "Calendar ID (default: primary)" }, eventId: { type: "string", description: "The event ID to delete" } }, required: ["eventId"] } } },
    { type: "function", function: { name: "list_calendars", description: "List all calendars accessible to the user.", parameters: { type: "object", properties: {} } } },
    { type: "function", function: { name: "create_calendar", description: "Create a new calendar.", parameters: { type: "object", properties: { summary: { type: "string", description: "Calendar name" }, description: { type: "string", description: "Calendar description" }, timeZone: { type: "string", description: "Time zone" } }, required: ["summary"] } } },
    { type: "function", function: { name: "quick_add_event", description: "Create event from natural language.", parameters: { type: "object", properties: { calendarId: { type: "string", description: "Calendar ID (default: primary)" }, text: { type: "string", description: "Natural language event description" } }, required: ["text"] } } },
    { type: "function", function: { name: "get_free_busy", description: "Check free/busy status.", parameters: { type: "object", properties: { timeMin: { type: "string", description: "Start of time range (ISO 8601)" }, timeMax: { type: "string", description: "End of time range (ISO 8601)" }, calendarIds: { type: "array", items: { type: "string" }, description: "Calendar IDs to check (default: ['primary'])" } }, required: ["timeMin", "timeMax"] } } },
    { type: "function", function: { name: "check_person_availability", description: "Check a person's availability.", parameters: { type: "object", properties: { person: { type: "string", description: "Person name or email to resolve from Gmail history" }, email: { type: "string", description: "Exact person email (preferred if known)" }, calendarId: { type: "string", description: "Calendar ID override (if known)" }, timeMin: { type: "string", description: "Start of time range (ISO 8601)" }, timeMax: { type: "string", description: "End of time range (ISO 8601)" }, durationMinutes: { type: "integer", description: "Minimum free slot length in minutes (default 30)" } }, required: ["timeMin", "timeMax"] } } },
    { type: "function", function: { name: "find_common_free_slots", description: "Find common free slots.", parameters: { type: "object", properties: { timeMin: { type: "string", description: "Start of time range (ISO 8601)" }, timeMax: { type: "string", description: "End of time range (ISO 8601)" }, people: { type: "array", items: { type: "string" }, description: "Names or emails to resolve and include" }, calendarIds: { type: "array", items: { type: "string" }, description: "Optional calendar IDs to include" }, includePrimary: { type: "boolean", description: "Include your primary calendar (default true)" }, durationMinutes: { type: "integer", description: "Minimum free slot length in minutes (default 30)" } }, required: ["timeMin", "timeMax"] } } },
    { type: "function", function: { name: "list_recurring_instances", description: "List recurring event instances.", parameters: { type: "object", properties: { calendarId: { type: "string", description: "Calendar ID (default: primary)" }, eventId: { type: "string", description: "The recurring event ID" }, maxResults: { type: "integer", description: "Max instances to return" }, timeMin: { type: "string", description: "Start time filter" }, timeMax: { type: "string", description: "End time filter" } }, required: ["eventId"] } } },
    { type: "function", function: { name: "move_event", description: "Move event to another calendar.", parameters: { type: "object", properties: { calendarId: { type: "string", description: "Source calendar ID (default: primary)" }, eventId: { type: "string", description: "The event ID to move" }, destinationCalendarId: { type: "string", description: "Destination calendar ID" } }, required: ["eventId", "destinationCalendarId"] } } },
    { type: "function", function: { name: "update_event_attendees", description: "Add or remove event attendees.", parameters: { type: "object", properties: { calendarId: { type: "string", description: "Calendar ID (default: primary)" }, eventId: { type: "string", description: "The event ID" }, addAttendees: { type: "array", items: { type: "string" }, description: "Email addresses to add (preferred parameter name)" }, attendees: { type: "array", items: { type: "string" }, description: "Alias for addAttendees - email addresses to add" }, removeAttendees: { type: "array", items: { type: "string" }, description: "Email addresses to remove" } }, required: ["eventId"] } } },
    { type: "function", function: { name: "get_calendar_colors", description: "Get calendar color options.", parameters: { type: "object", properties: {} } } },
    { type: "function", function: { name: "clear_calendar", description: "Clear all events from a calendar.", parameters: { type: "object", properties: { calendarId: { type: "string", description: "The calendar ID to clear (cannot be primary)" } }, required: ["calendarId"] } } },
    { type: "function", function: { name: "watch_events", description: "Watch calendar for changes.", parameters: { type: "object", properties: { calendarId: { type: "string", description: "Calendar ID (default: primary)" }, webhookUrl: { type: "string", description: "Public webhook URL to receive notifications" } }, required: ["webhookUrl"] } } }
];

const gchatTools = [
    { type: "function", function: { name: "list_chat_spaces", description: "List Chat spaces.", parameters: { type: "object", properties: { maxResults: { type: "integer", description: "Max spaces to return (default 20)" } } } } },
    { type: "function", function: { name: "send_chat_message", description: "Send a Chat message.", parameters: { type: "object", properties: { spaceId: { type: "string", description: "Space ID or full name like spaces/AAAA..." }, text: { type: "string", description: "Message text to send" } }, required: ["spaceId", "text"] } } },
    { type: "function", function: { name: "list_chat_messages", description: "List Chat messages.", parameters: { type: "object", properties: { spaceId: { type: "string", description: "Space ID or full name like spaces/AAAA..." }, maxResults: { type: "integer", description: "Max messages to return (default 20)" } }, required: ["spaceId"] } } }
];

const driveTools = [
    { type: "function", function: { name: "list_drive_files", description: "List Drive files/folders.", parameters: { type: "object", properties: { query: { type: "string", description: "Drive query string, e.g. name contains 'Q1' and mimeType contains 'spreadsheet'" }, pageSize: { type: "integer", description: "Max files to return (default 25)" }, orderBy: { type: "string", description: "Sort order (default 'modifiedTime desc')" }, includeTrashed: { type: "boolean", description: "Include trashed files (default false)" } } } } },
    { type: "function", function: { name: "get_drive_file", description: "Get Drive file metadata.", parameters: { type: "object", properties: { fileId: { type: "string", description: "Drive file ID" } }, required: ["fileId"] } } },
    { type: "function", function: { name: "create_drive_folder", description: "Create a Drive folder.", parameters: { type: "object", properties: { name: { type: "string", description: "Folder name" }, parentId: { type: "string", description: "Optional parent folder ID" } }, required: ["name"] } } },
    { type: "function", function: { name: "create_drive_file", description: "Create a file in Drive.", parameters: { type: "object", properties: { name: { type: "string", description: "File name" }, content: { type: "string", description: "Text content (use this OR localPath, not both)" }, mimeType: { type: "string", description: "MIME type (auto-detected from file extension if using localPath)" }, parentId: { type: "string", description: "Optional parent folder ID" }, localPath: { type: "string", description: "Local server file path (from attached files). Use this to upload user's attached files to Drive." } }, required: ["name"] } } },
    { type: "function", function: { name: "update_drive_file", description: "Update a Drive file.", parameters: { type: "object", properties: { fileId: { type: "string", description: "Drive file ID" }, content: { type: "string", description: "New text content" }, name: { type: "string", description: "New file name" }, mimeType: { type: "string", description: "MIME type for content uploads (default text/plain)" } }, required: ["fileId"] } } },
    { type: "function", function: { name: "delete_drive_file", description: "Delete a Drive file.", parameters: { type: "object", properties: { fileId: { type: "string", description: "Drive file ID" }, permanent: { type: "boolean", description: "If true, permanently delete. Otherwise move to trash." } }, required: ["fileId"] } } },
    { type: "function", function: { name: "copy_drive_file", description: "Copy a Drive file.", parameters: { type: "object", properties: { fileId: { type: "string", description: "Source Drive file ID" }, name: { type: "string", description: "Optional new file name" }, parentId: { type: "string", description: "Optional parent folder for copy" } }, required: ["fileId"] } } },
    { type: "function", function: { name: "move_drive_file", description: "Move a Drive file.", parameters: { type: "object", properties: { fileId: { type: "string", description: "Drive file ID" }, newParentId: { type: "string", description: "Destination folder ID" } }, required: ["fileId", "newParentId"] } } },
    { type: "function", function: { name: "share_drive_file", description: "Share a Drive file.", parameters: { type: "object", properties: { fileId: { type: "string", description: "Drive file ID" }, emailAddress: { type: "string", description: "User email address to share with" }, role: { type: "string", description: "Permission role: reader, commenter, writer, organizer, fileOrganizer (default reader)" }, sendNotificationEmail: { type: "boolean", description: "Send share email notification (default true)" } }, required: ["fileId", "emailAddress"] } } },
    { type: "function", function: { name: "download_drive_file", description: "Get browser download link for a Drive file.", parameters: { type: "object", properties: { fileId: { type: "string", description: "Drive file ID" }, format: { type: "string", description: "Optional export format for Google Workspace files (e.g. docx, pdf, txt, xlsx, csv)." }, filename: { type: "string", description: "Optional custom downloaded file name." } }, required: ["fileId"] } } },
    { type: "function", function: { name: "download_drive_file_to_local", description: "Download a Drive file to server disk. Returns localPath for attachments.", parameters: { type: "object", properties: { fileId: { type: "string", description: "Drive file ID" }, format: { type: "string", description: "Export format for Google Workspace files (pdf, docx, xlsx, csv, pptx, txt)" } }, required: ["fileId"] } } },
    { type: "function", function: { name: "extract_drive_file_text", description: "Extract text from a Drive file.", parameters: { type: "object", properties: { fileId: { type: "string", description: "Drive file ID" }, maxBytes: { type: "integer", description: "Maximum text bytes to return (default 40000, max 120000)." } }, required: ["fileId"] } } },
    { type: "function", function: { name: "append_drive_document_text", description: "Append text to a Drive document.", parameters: { type: "object", properties: { fileId: { type: "string", description: "Drive file ID (Google Doc or text-like file)." }, text: { type: "string", description: "Text to append to the end of the document." } }, required: ["fileId", "text"] } } },
    { type: "function", function: { name: "convert_file_to_google_doc", description: "Convert a file to Google Doc.", parameters: { type: "object", properties: { fileId: { type: "string", description: "Source Drive file ID (e.g. a PDF)." }, title: { type: "string", description: "Optional title for the converted Google Doc." }, parentId: { type: "string", description: "Optional destination Drive folder ID for the converted doc." }, downloadConverted: { type: "boolean", description: "If true, also return a download URL for the converted document." }, downloadFormat: { type: "string", description: "Optional export format for converted doc (docx, pdf, txt). Default docx." } }, required: ["fileId"] } } },
    { type: "function", function: { name: "convert_file_to_google_sheet", description: "Convert a file to Google Sheet.", parameters: { type: "object", properties: { fileId: { type: "string", description: "Source Drive file ID (e.g. CSV/XLSX)." }, title: { type: "string", description: "Optional title for the converted Google Sheet." }, parentId: { type: "string", description: "Optional destination Drive folder ID for the converted sheet." }, downloadConverted: { type: "boolean", description: "If true, also return a download URL for the converted spreadsheet." }, downloadFormat: { type: "string", description: "Optional export format for converted sheet (xlsx, csv, pdf). Default xlsx." } }, required: ["fileId"] } } }
];

const sheetsTools = [
    { type: "function", function: { name: "list_spreadsheets", description: "List Google Sheets.", parameters: { type: "object", properties: { query: { type: "string", description: "Optional Drive query filter" }, maxResults: { type: "integer", description: "Max spreadsheets to return (default 25)" } } } } },
    { type: "function", function: { name: "create_spreadsheet", description: "Create a new Google Spreadsheet.", parameters: { type: "object", properties: { title: { type: "string", description: "Spreadsheet title" }, sheets: { type: "array", items: { type: "string" }, description: "Optional sheet tab titles" } }, required: ["title"] } } },
    { type: "function", function: { name: "get_spreadsheet", description: "Get spreadsheet metadata.", parameters: { type: "object", properties: { spreadsheetId: { type: "string", description: "Spreadsheet ID" }, includeGridData: { type: "boolean", description: "Include cell grid data (default false)" } }, required: ["spreadsheetId"] } } },
    { type: "function", function: { name: "list_sheet_tabs", description: "List sheet tabs.", parameters: { type: "object", properties: { spreadsheetId: { type: "string", description: "Spreadsheet ID" } }, required: ["spreadsheetId"] } } },
    { type: "function", function: { name: "add_sheet_tab", description: "Add a sheet tab.", parameters: { type: "object", properties: { spreadsheetId: { type: "string", description: "Spreadsheet ID" }, title: { type: "string", description: "New tab title" }, rows: { type: "integer", description: "Initial row count (default 1000)" }, columns: { type: "integer", description: "Initial column count (default 26)" } }, required: ["spreadsheetId", "title"] } } },
    { type: "function", function: { name: "delete_sheet_tab", description: "Delete a sheet tab.", parameters: { type: "object", properties: { spreadsheetId: { type: "string", description: "Spreadsheet ID" }, sheetId: { type: "integer", description: "Numeric sheet ID" } }, required: ["spreadsheetId", "sheetId"] } } },
    { type: "function", function: { name: "read_sheet_values", description: "Read spreadsheet values.", parameters: { type: "object", properties: { spreadsheetId: { type: "string", description: "Spreadsheet ID" }, range: { type: "string", description: "A1 range (e.g. Sheet1!A1:D20)" }, valueRenderOption: { type: "string", description: "FORMATTED_VALUE, UNFORMATTED_VALUE, or FORMULA" }, dateTimeRenderOption: { type: "string", description: "SERIAL_NUMBER or FORMATTED_STRING" } }, required: ["spreadsheetId", "range"] } } },
    { type: "function", function: { name: "update_sheet_values", description: "Update spreadsheet values.", parameters: { type: "object", properties: { spreadsheetId: { type: "string", description: "Spreadsheet ID" }, range: { type: "string", description: "A1 range to write" }, values: { type: "array", description: "2D array of rows, e.g. [[\"Name\",\"Role\"],[\"Rishi\",\"Lead\"]]", items: { type: "array", items: { type: "string" } } }, valueInputOption: { type: "string", description: "RAW or USER_ENTERED (default USER_ENTERED)" }, majorDimension: { type: "string", description: "ROWS or COLUMNS (default ROWS)" } }, required: ["spreadsheetId", "range", "values"] } } },
    { type: "function", function: { name: "update_timesheet_hours", description: "Update timesheet row by date.", parameters: { type: "object", properties: { spreadsheetId: { type: "string", description: "Spreadsheet ID" }, sheetName: { type: "string", description: "Tab name (default Tracker)" }, date: { type: "string", description: "Date to match, e.g. 6-Feb-2026 or 2026-02-06" }, billingHours: { type: "number", description: "Billing hours value to set" }, taskDetails: { type: "string", description: "Task details/description to set" }, nonBillingHours: { type: "number", description: "Non-billing hours value to set" }, projectName: { type: "string", description: "Project name to set" }, moduleName: { type: "string", description: "Module name to set" }, month: { type: "string", description: "Month label to set (e.g. February 2026)" }, dateColumn: { type: "string", description: "Date column letter (default B)" }, taskDetailsColumn: { type: "string", description: "Task Details column letter (default C)" }, billingHoursColumn: { type: "string", description: "Billing hours column letter (default D)" }, nonBillingHoursColumn: { type: "string", description: "Non-billing hours column letter (default E)" }, projectNameColumn: { type: "string", description: "Project name column letter (default F)" }, moduleNameColumn: { type: "string", description: "Module name column letter (default G)" }, monthColumn: { type: "string", description: "Month column letter (default A)" }, searchRange: { type: "string", description: "A1 range to search (default A1:Z3000)" }, preferEmptyBilling: { type: "boolean", description: "Prefer row where billing cell is empty when duplicates exist and billingHours is being updated (default true)" } }, required: ["spreadsheetId", "date"] } } },
    { type: "function", function: { name: "append_sheet_values", description: "Append rows to a spreadsheet.", parameters: { type: "object", properties: { spreadsheetId: { type: "string", description: "Spreadsheet ID" }, range: { type: "string", description: "A1 target range (e.g. Sheet1!A:D)" }, values: { type: "array", description: "2D array of rows to append", items: { type: "array", items: { type: "string" } } }, valueInputOption: { type: "string", description: "RAW or USER_ENTERED (default USER_ENTERED)" }, insertDataOption: { type: "string", description: "INSERT_ROWS or OVERWRITE (default INSERT_ROWS)" } }, required: ["spreadsheetId", "range", "values"] } } },
    { type: "function", function: { name: "clear_sheet_values", description: "Clear spreadsheet values.", parameters: { type: "object", properties: { spreadsheetId: { type: "string", description: "Spreadsheet ID" }, range: { type: "string", description: "A1 range to clear" } }, required: ["spreadsheetId", "range"] } } }
];

const githubTools = [
    { type: "function", function: { name: "list_repos", description: "List GitHub repositories.", parameters: { type: "object", properties: { username: { type: "string", description: "GitHub username (omit for your own repos)" }, sort: { type: "string", description: "Sort by: created, updated, pushed, full_name (default: updated)" }, perPage: { type: "integer", description: "Results per page (default 30)" } } } } },
    { type: "function", function: { name: "get_repo", description: "Get repository details.", parameters: { type: "object", properties: { owner: { type: "string", description: "Repository owner" }, repo: { type: "string", description: "Repository name" } }, required: ["owner", "repo"] } } },
    { type: "function", function: { name: "create_repo", description: "Create a new GitHub repository.", parameters: { type: "object", properties: { name: { type: "string", description: "Repository name" }, description: { type: "string", description: "Repository description" }, isPrivate: { type: "boolean", description: "Make repository private (default: false)" }, autoInit: { type: "boolean", description: "Initialize with README (default: true)" } }, required: ["name"] } } },
    { type: "function", function: { name: "list_issues", description: "List issues for a repository.", parameters: { type: "object", properties: { owner: { type: "string", description: "Repository owner" }, repo: { type: "string", description: "Repository name" }, state: { type: "string", description: "Issue state: open, closed, all (default: open)" }, labels: { type: "string", description: "Comma-separated label names" }, perPage: { type: "integer", description: "Results per page (default 30)" } }, required: ["owner", "repo"] } } },
    { type: "function", function: { name: "create_issue", description: "Create a new issue in a repository.", parameters: { type: "object", properties: { owner: { type: "string", description: "Repository owner" }, repo: { type: "string", description: "Repository name" }, title: { type: "string", description: "Issue title" }, body: { type: "string", description: "Issue body (markdown)" }, labels: { type: "array", items: { type: "string" }, description: "Label names" }, assignees: { type: "array", items: { type: "string" }, description: "Assignee usernames" } }, required: ["owner", "repo", "title"] } } },
    { type: "function", function: { name: "update_issue", description: "Update an issue.", parameters: { type: "object", properties: { owner: { type: "string", description: "Repository owner" }, repo: { type: "string", description: "Repository name" }, issueNumber: { type: "integer", description: "Issue number" }, title: { type: "string", description: "New title" }, body: { type: "string", description: "New body" }, state: { type: "string", description: "New state: open or closed" }, labels: { type: "array", items: { type: "string" }, description: "Labels to set" }, assignees: { type: "array", items: { type: "string" }, description: "Assignees to set" } }, required: ["owner", "repo", "issueNumber"] } } },
    { type: "function", function: { name: "list_pull_requests", description: "List pull requests for a repository.", parameters: { type: "object", properties: { owner: { type: "string", description: "Repository owner" }, repo: { type: "string", description: "Repository name" }, state: { type: "string", description: "PR state: open, closed, all (default: open)" }, perPage: { type: "integer", description: "Results per page (default 30)" } }, required: ["owner", "repo"] } } },
    { type: "function", function: { name: "get_pull_request", description: "Get pull request details.", parameters: { type: "object", properties: { owner: { type: "string", description: "Repository owner" }, repo: { type: "string", description: "Repository name" }, pullNumber: { type: "integer", description: "Pull request number" } }, required: ["owner", "repo", "pullNumber"] } } },
    { type: "function", function: { name: "create_pull_request", description: "Create a new pull request.", parameters: { type: "object", properties: { owner: { type: "string", description: "Repository owner" }, repo: { type: "string", description: "Repository name" }, title: { type: "string", description: "PR title" }, body: { type: "string", description: "PR description" }, head: { type: "string", description: "Branch containing changes" }, base: { type: "string", description: "Branch to merge into" } }, required: ["owner", "repo", "title", "head", "base"] } } },
    { type: "function", function: { name: "merge_pull_request", description: "Merge a pull request.", parameters: { type: "object", properties: { owner: { type: "string", description: "Repository owner" }, repo: { type: "string", description: "Repository name" }, pullNumber: { type: "integer", description: "Pull request number" }, mergeMethod: { type: "string", description: "Merge method: merge, squash, rebase (default: merge)" }, commitMessage: { type: "string", description: "Custom merge commit message" } }, required: ["owner", "repo", "pullNumber"] } } },
    { type: "function", function: { name: "list_branches", description: "List branches in a repository.", parameters: { type: "object", properties: { owner: { type: "string", description: "Repository owner" }, repo: { type: "string", description: "Repository name" }, perPage: { type: "integer", description: "Results per page (default 30)" } }, required: ["owner", "repo"] } } },
    { type: "function", function: { name: "create_branch", description: "Create a new branch.", parameters: { type: "object", properties: { owner: { type: "string", description: "Repository owner" }, repo: { type: "string", description: "Repository name" }, branchName: { type: "string", description: "New branch name" }, fromBranch: { type: "string", description: "Source branch (default: main)" } }, required: ["owner", "repo", "branchName"] } } },
    { type: "function", function: { name: "get_file_content", description: "Get file content from a repo.", parameters: { type: "object", properties: { owner: { type: "string", description: "Repository owner" }, repo: { type: "string", description: "Repository name" }, filePath: { type: "string", description: "Path to the file" }, ref: { type: "string", description: "Branch or commit SHA (default: default branch)" } }, required: ["owner", "repo", "filePath"] } } },
    { type: "function", function: { name: "create_or_update_file", description: "Create or update a file in a repo.", parameters: { type: "object", properties: { owner: { type: "string", description: "Repository owner" }, repo: { type: "string", description: "Repository name" }, filePath: { type: "string", description: "Path for the file" }, content: { type: "string", description: "File content" }, message: { type: "string", description: "Commit message" }, branch: { type: "string", description: "Target branch" }, sha: { type: "string", description: "SHA of file being replaced (required for updates)" } }, required: ["owner", "repo", "filePath", "content", "message"] } } },
    { type: "function", function: { name: "search_repos", description: "Search GitHub repositories.", parameters: { type: "object", properties: { query: { type: "string", description: "Search query (e.g. 'react language:javascript stars:>1000')" }, sort: { type: "string", description: "Sort by: stars, forks, updated (default: stars)" }, perPage: { type: "integer", description: "Results per page (default 20)" } }, required: ["query"] } } },
    { type: "function", function: { name: "search_code", description: "Search code on GitHub.", parameters: { type: "object", properties: { query: { type: "string", description: "Code search query (e.g. 'useState repo:facebook/react')" }, perPage: { type: "integer", description: "Results per page (default 20)" } }, required: ["query"] } } },
    { type: "function", function: { name: "list_commits", description: "List recent commits.", parameters: { type: "object", properties: { owner: { type: "string", description: "Repository owner" }, repo: { type: "string", description: "Repository name" }, sha: { type: "string", description: "Branch name or commit SHA" }, perPage: { type: "integer", description: "Results per page (default 20)" } }, required: ["owner", "repo"] } } },
    { type: "function", function: { name: "revert_commit", description: "Revert a commit.", parameters: { type: "object", properties: { owner: { type: "string", description: "Repository owner" }, repo: { type: "string", description: "Repository name" }, commitSha: { type: "string", description: "Full or short SHA of the commit to revert" }, branch: { type: "string", description: "Branch to revert on (default: main)" } }, required: ["owner", "repo", "commitSha"] } } },
    { type: "function", function: { name: "reset_branch", description: "Hard reset a branch. use 'removeCommitSha' to remove a specific recent commit (and everything after it), OR 'targetSha' to reset the branch to a specific past state.", parameters: { type: "object", properties: { owner: { type: "string", description: "Repository owner" }, repo: { type: "string", description: "Repository name" }, branch: { type: "string", description: "Branch to reset" }, targetSha: { type: "string", description: "Commit SHA to reset TO (retain history up to here)" }, removeCommitSha: { type: "string", description: "Commit SHA to REMOVE (resets to its parent)" } }, required: ["owner", "repo", "branch"] } } },
    { type: "function", function: { name: "get_user_profile", description: "Get GitHub user profile.", parameters: { type: "object", properties: { username: { type: "string", description: "GitHub username (omit for your own)" } } } } },
    { type: "function", function: { name: "list_notifications", description: "List your GitHub notifications.", parameters: { type: "object", properties: { all: { type: "boolean", description: "Show all including read (default: false)" }, perPage: { type: "integer", description: "Results per page (default 20)" } } } } },
    { type: "function", function: { name: "list_gists", description: "List your GitHub gists.", parameters: { type: "object", properties: { perPage: { type: "integer", description: "Results per page (default 20)" } } } } }
];

const outlookTools = [
    { type: "function", function: { name: "outlook_send_email", description: "Send an email via Outlook.", parameters: { type: "object", properties: { to: { type: "array", items: { type: "string" }, description: "Recipient email addresses" }, subject: { type: "string", description: "Email subject line" }, body: { type: "string", description: "Email body (plain text or HTML)" }, confirmSend: { type: "boolean", description: "Set true after user confirms." }, cc: { type: "array", items: { type: "string" }, description: "CC addresses" }, bcc: { type: "array", items: { type: "string" }, description: "BCC addresses" } }, required: ["to", "subject", "body"] } } },
    { type: "function", function: { name: "outlook_list_emails", description: "List Outlook emails.", parameters: { type: "object", properties: { maxResults: { type: "integer", description: "Number of emails to return (default 20)" }, folder: { type: "string", description: "Folder name: inbox, sentitems, drafts, deleteditems, junkemail (default: inbox)" } } } } },
    { type: "function", function: { name: "outlook_read_email", description: "Read an Outlook email.", parameters: { type: "object", properties: { messageId: { type: "string", description: "Outlook message ID" } }, required: ["messageId"] } } },
    { type: "function", function: { name: "outlook_search_emails", description: "Search Outlook emails (KQL: from:, subject:, hasAttachment:true).", parameters: { type: "object", properties: { query: { type: "string", description: "Search query (e.g., 'from:john subject:meeting')" }, maxResults: { type: "integer", description: "Number of results (default 20)" } }, required: ["query"] } } },
    { type: "function", function: { name: "outlook_reply_to_email", description: "Reply to an Outlook email.", parameters: { type: "object", properties: { messageId: { type: "string", description: "Message ID to reply to" }, body: { type: "string", description: "Reply body (HTML or text)" }, confirmSend: { type: "boolean", description: "Set true after user confirms." } }, required: ["messageId", "body"] } } },
    { type: "function", function: { name: "outlook_forward_email", description: "Forward an Outlook email to new recipients.", parameters: { type: "object", properties: { messageId: { type: "string", description: "Message ID to forward" }, to: { type: "array", items: { type: "string" }, description: "Forward recipient addresses" }, comment: { type: "string", description: "Comment to include" }, confirmSend: { type: "boolean", description: "Set true after user confirms." } }, required: ["messageId", "to"] } } },
    { type: "function", function: { name: "outlook_delete_email", description: "Delete an Outlook email.", parameters: { type: "object", properties: { messageId: { type: "string", description: "Message ID to delete" } }, required: ["messageId"] } } },
    { type: "function", function: { name: "outlook_move_email", description: "Move an Outlook email.", parameters: { type: "object", properties: { messageId: { type: "string", description: "Message ID to move" }, destinationFolderId: { type: "string", description: "Destination folder ID (use outlook_list_folders to find IDs)" } }, required: ["messageId", "destinationFolderId"] } } },
    { type: "function", function: { name: "outlook_mark_as_read", description: "Mark an Outlook email as read.", parameters: { type: "object", properties: { messageId: { type: "string", description: "Message ID" } }, required: ["messageId"] } } },
    { type: "function", function: { name: "outlook_mark_as_unread", description: "Mark an Outlook email as unread.", parameters: { type: "object", properties: { messageId: { type: "string", description: "Message ID" } }, required: ["messageId"] } } },
    { type: "function", function: { name: "outlook_list_folders", description: "List Outlook folders.", parameters: { type: "object", properties: {} } } },
    { type: "function", function: { name: "outlook_create_folder", description: "Create an Outlook folder.", parameters: { type: "object", properties: { name: { type: "string", description: "Folder display name" }, parentFolderId: { type: "string", description: "Parent folder ID (omit for top-level)" } }, required: ["name"] } } },
    { type: "function", function: { name: "outlook_get_attachments", description: "Get Outlook attachments.", parameters: { type: "object", properties: { messageId: { type: "string", description: "Message ID" } }, required: ["messageId"] } } },
    { type: "function", function: { name: "outlook_create_draft", description: "Create an Outlook draft.", parameters: { type: "object", properties: { to: { type: "array", items: { type: "string" }, description: "Recipient addresses" }, subject: { type: "string", description: "Subject line" }, body: { type: "string", description: "Email body (HTML or text)" }, cc: { type: "array", items: { type: "string" }, description: "CC addresses" }, bcc: { type: "array", items: { type: "string" }, description: "BCC addresses" } } } } },
    { type: "function", function: { name: "outlook_send_draft", description: "Send an Outlook draft.", parameters: { type: "object", properties: { messageId: { type: "string", description: "Draft message ID" }, confirmSend: { type: "boolean", description: "Set true after user confirms." } }, required: ["messageId"] } } },
    { type: "function", function: { name: "outlook_list_drafts", description: "List Outlook drafts.", parameters: { type: "object", properties: { maxResults: { type: "integer", description: "Number of drafts to return (default 20)" } } } } },
    { type: "function", function: { name: "outlook_flag_email", description: "Flag an Outlook email.", parameters: { type: "object", properties: { messageId: { type: "string", description: "Message ID" }, flagStatus: { type: "string", description: "Flag status: flagged, complete, notFlagged (default: flagged)" } }, required: ["messageId"] } } },
    { type: "function", function: { name: "outlook_get_user_profile", description: "Get Outlook user profile.", parameters: { type: "object", properties: {} } } }
];

const docsTools = [
    { type: "function", function: { name: "list_documents", description: "List Google Docs.", parameters: { type: "object", properties: { query: { type: "string", description: "Optional name filter (e.g. 'Meeting Notes')" }, pageSize: { type: "integer", description: "Max documents to return (default 25)" } } } } },
    { type: "function", function: { name: "get_document", description: "Get Doc metadata.", parameters: { type: "object", properties: { documentId: { type: "string", description: "Google Doc document ID" } }, required: ["documentId"] } } },
    { type: "function", function: { name: "create_document", description: "Create a new Google Doc.", parameters: { type: "object", properties: { title: { type: "string", description: "Document title" }, content: { type: "string", description: "Optional initial text content" } }, required: ["title"] } } },
    { type: "function", function: { name: "insert_text", description: "Insert text in a Doc.", parameters: { type: "object", properties: { documentId: { type: "string", description: "Google Doc document ID" }, text: { type: "string", description: "Text to insert" }, index: { type: "integer", description: "Character index to insert at (default 1 = start)" } }, required: ["documentId", "text"] } } },
    { type: "function", function: { name: "replace_text", description: "Find and replace text in a Doc.", parameters: { type: "object", properties: { documentId: { type: "string", description: "Google Doc document ID" }, findText: { type: "string", description: "Text to find" }, replaceWith: { type: "string", description: "Replacement text" }, matchCase: { type: "boolean", description: "Case-sensitive match (default false)" } }, required: ["documentId", "findText", "replaceWith"] } } },
    { type: "function", function: { name: "delete_content", description: "Delete content in a Doc.", parameters: { type: "object", properties: { documentId: { type: "string", description: "Google Doc document ID" }, startIndex: { type: "integer", description: "Start character index" }, endIndex: { type: "integer", description: "End character index" } }, required: ["documentId", "startIndex", "endIndex"] } } },
    { type: "function", function: { name: "append_text", description: "Append text to a Doc.", parameters: { type: "object", properties: { documentId: { type: "string", description: "Google Doc document ID" }, text: { type: "string", description: "Text to append" } }, required: ["documentId", "text"] } } },
    { type: "function", function: { name: "get_document_text", description: "Get Doc text content.", parameters: { type: "object", properties: { documentId: { type: "string", description: "Google Doc document ID" } }, required: ["documentId"] } } }
];

const meetingTranscriptionTools = [
    {
        type: "function",
        function: {
            name: "list_meeting_transcriptions",
            description: "Search Google Drive for meeting transcript documents (defaults to Shared with me) and return file name + date + link.",
            parameters: {
                type: "object",
                properties: {
                    query: { type: "string", description: "Optional keyword or Drive query (e.g. team sync, 'name contains \\'retrospective\\'')" },
                    date: { type: "string", description: "Optional exact date filter in YYYY-MM-DD." },
                    pageSize: { type: "integer", description: "Max files to return (default 20)." },
                    sharedWithMe: { type: "boolean", description: "Search only files shared with me (default true)." }
                }
            }
        }
    },
    {
        type: "function",
        function: {
            name: "open_meeting_transcription_file",
            description: "Get metadata and web link for a specific meeting transcription file.",
            parameters: {
                type: "object",
                properties: {
                    fileId: { type: "string", description: "Google Drive file ID." }
                },
                required: ["fileId"]
            }
        }
    },
    {
        type: "function",
        function: {
            name: "summarize_meeting_transcription",
            description: "Read a meeting transcription document and generate Summary, Action Items, Next Steps, plus source file link.",
            parameters: {
                type: "object",
                properties: {
                    fileId: { type: "string", description: "Google Drive file ID of the transcript document." }
                },
                required: ["fileId"]
            }
        }
    }
];

const teamsTools = [
    { type: "function", function: { name: "teams_list_teams", description: "List Teams.", parameters: { type: "object", properties: {} } } },
    { type: "function", function: { name: "teams_get_team", description: "Get Team details.", parameters: { type: "object", properties: { teamId: { type: "string", description: "Team ID" } }, required: ["teamId"] } } },
    { type: "function", function: { name: "teams_list_channels", description: "List Team channels.", parameters: { type: "object", properties: { teamId: { type: "string", description: "Team ID" } }, required: ["teamId"] } } },
    { type: "function", function: { name: "teams_send_channel_message", description: "Send a channel message.", parameters: { type: "object", properties: { teamId: { type: "string", description: "Team ID" }, channelId: { type: "string", description: "Channel ID" }, message: { type: "string", description: "Message content (plain text or HTML)" }, contentType: { type: "string", description: "Content type: text or html (default text)" } }, required: ["teamId", "channelId", "message"] } } },
    { type: "function", function: { name: "teams_list_channel_messages", description: "List channel messages.", parameters: { type: "object", properties: { teamId: { type: "string", description: "Team ID" }, channelId: { type: "string", description: "Channel ID" }, top: { type: "integer", description: "Number of messages to return (default 20)" } }, required: ["teamId", "channelId"] } } },
    { type: "function", function: { name: "teams_list_chats", description: "List Teams chats.", parameters: { type: "object", properties: { top: { type: "integer", description: "Number of chats to return (default 20)" } } } } },
    { type: "function", function: { name: "teams_send_chat_message", description: "Send a Teams chat message.", parameters: { type: "object", properties: { chatId: { type: "string", description: "Chat ID" }, message: { type: "string", description: "Message content (plain text or HTML)" }, contentType: { type: "string", description: "Content type: text or html (default text)" } }, required: ["chatId", "message"] } } },
    { type: "function", function: { name: "teams_list_chat_messages", description: "List Teams chat messages.", parameters: { type: "object", properties: { chatId: { type: "string", description: "Chat ID" }, top: { type: "integer", description: "Number of messages to return (default 20)" } }, required: ["chatId"] } } },
    { type: "function", function: { name: "teams_create_chat", description: "Create a Teams chat.", parameters: { type: "object", properties: { chatType: { type: "string", description: "Chat type: oneOnOne or group" }, members: { type: "array", items: { type: "string" }, description: "Array of member email addresses" }, topic: { type: "string", description: "Chat topic (for group chats)" } }, required: ["chatType", "members"] } } },
    { type: "function", function: { name: "teams_get_chat_members", description: "List chat members.", parameters: { type: "object", properties: { chatId: { type: "string", description: "Chat ID" } }, required: ["chatId"] } } }
];

const gcsTools = [
    { type: "function", function: { name: "gcs_list_buckets", description: "List GCS buckets.", parameters: { type: "object", properties: {} } } },
    { type: "function", function: { name: "gcs_get_bucket", description: "Get bucket metadata.", parameters: { type: "object", properties: { bucket: { type: "string", description: "Bucket name" } }, required: ["bucket"] } } },
    { type: "function", function: { name: "gcs_create_bucket", description: "Create a GCS bucket.", parameters: { type: "object", properties: { name: { type: "string", description: "Bucket name (globally unique)" }, location: { type: "string", description: "Bucket location (default US)" }, storageClass: { type: "string", description: "Storage class: STANDARD, NEARLINE, COLDLINE, ARCHIVE (default STANDARD)" } }, required: ["name"] } } },
    { type: "function", function: { name: "gcs_delete_bucket", description: "Delete a GCS bucket.", parameters: { type: "object", properties: { bucket: { type: "string", description: "Bucket name" } }, required: ["bucket"] } } },
    { type: "function", function: { name: "gcs_list_objects", description: "List GCS objects.", parameters: { type: "object", properties: { bucket: { type: "string", description: "Bucket name" }, prefix: { type: "string", description: "Filter objects by prefix (folder path)" }, maxResults: { type: "integer", description: "Maximum number of objects to return" } }, required: ["bucket"] } } },
    { type: "function", function: { name: "gcs_upload_object", description: "Upload to GCS.", parameters: { type: "object", properties: { bucket: { type: "string", description: "Bucket name" }, name: { type: "string", description: "Object name (path in bucket)" }, content: { type: "string", description: "Text content to upload (use this OR localPath, not both)" }, contentType: { type: "string", description: "MIME type (auto-detected from file extension if using localPath)" }, localPath: { type: "string", description: "Local server file path (from attached files). Use this to upload user's attached files to GCS." } }, required: ["bucket", "name"] } } },
    { type: "function", function: { name: "gcs_download_object", description: "Download from GCS.", parameters: { type: "object", properties: { bucket: { type: "string", description: "Bucket name" }, name: { type: "string", description: "Object name (path in bucket)" }, filename: { type: "string", description: "Optional: desired filename for download (defaults to object name)" } }, required: ["bucket", "name"] } } },
    { type: "function", function: { name: "gcs_delete_object", description: "Delete a GCS object.", parameters: { type: "object", properties: { bucket: { type: "string", description: "Bucket name" }, name: { type: "string", description: "Object name (path in bucket)" } }, required: ["bucket", "name"] } } },
    { type: "function", function: { name: "gcs_copy_object", description: "Copy a GCS object.", parameters: { type: "object", properties: { sourceBucket: { type: "string", description: "Source bucket name" }, sourceObject: { type: "string", description: "Source object name" }, destBucket: { type: "string", description: "Destination bucket name (default same as source)" }, destObject: { type: "string", description: "Destination object name (default same as source)" } }, required: ["sourceBucket", "sourceObject"] } } },
    { type: "function", function: { name: "gcs_get_object_metadata", description: "Get object metadata.", parameters: { type: "object", properties: { bucket: { type: "string", description: "Bucket name" }, name: { type: "string", description: "Object name (path in bucket)" } }, required: ["bucket", "name"] } } },
    { type: "function", function: { name: "gcs_move_object", description: "Move a GCS object.", parameters: { type: "object", properties: { sourceBucket: { type: "string", description: "Source bucket name" }, sourceObject: { type: "string", description: "Source object name" }, destBucket: { type: "string", description: "Destination bucket name (defaults to source bucket)" }, destObject: { type: "string", description: "Destination object name (required for move)" } }, required: ["sourceBucket", "sourceObject", "destObject"] } } },
    { type: "function", function: { name: "gcs_rename_object", description: "Rename a GCS object.", parameters: { type: "object", properties: { bucket: { type: "string", description: "Bucket name" }, oldName: { type: "string", description: "Current object name" }, newName: { type: "string", description: "New object name" } }, required: ["bucket", "oldName", "newName"] } } },
    { type: "function", function: { name: "gcs_make_object_public", description: "Make object public.", parameters: { type: "object", properties: { bucket: { type: "string", description: "Bucket name" }, name: { type: "string", description: "Object name" } }, required: ["bucket", "name"] } } },
    { type: "function", function: { name: "gcs_make_object_private", description: "Make object private.", parameters: { type: "object", properties: { bucket: { type: "string", description: "Bucket name" }, name: { type: "string", description: "Object name" } }, required: ["bucket", "name"] } } },
    { type: "function", function: { name: "gcs_generate_signed_url", description: "Generate signed URL.", parameters: { type: "object", properties: { bucket: { type: "string", description: "Bucket name" }, name: { type: "string", description: "Object name" }, expirationHours: { type: "integer", description: "URL validity in hours (default 24)" } }, required: ["bucket", "name"] } } },
    { type: "function", function: { name: "gcs_batch_delete_objects", description: "Batch delete objects by prefix.", parameters: { type: "object", properties: { bucket: { type: "string", description: "Bucket name" }, prefix: { type: "string", description: "Delete all objects with this prefix (e.g., 'old-files/' or 'temp-')" }, confirmDelete: { type: "boolean", description: "MUST be true to confirm deletion" } }, required: ["bucket", "prefix", "confirmDelete"] } } }
];

// ============================================================
//  TOOL EXECUTION ROUTERS
// ============================================================

// Gmail tool names for fast lookup
const gmailToolNames = new Set(gmailTools.map(t => t.function.name));
const calendarToolNames = new Set(calendarTools.map(t => t.function.name));
const gchatToolNames = new Set(gchatTools.map(t => t.function.name));
const driveToolNames = new Set(driveTools.map(t => t.function.name));
const sheetsToolNames = new Set(sheetsTools.map(t => t.function.name));
const githubToolNames = new Set(githubTools.map(t => t.function.name));
const outlookToolNames = new Set(outlookTools.map(t => t.function.name));
const docsToolNames = new Set(docsTools.map(t => t.function.name));
const meetingTranscriptionToolNames = new Set(meetingTranscriptionTools.map(t => t.function.name));
const teamsToolNames = new Set(teamsTools.map(t => t.function.name));
const gcsToolNames = new Set(gcsTools.map(t => t.function.name));
const docsListOnlyTools = docsTools.filter(tool => tool.function.name === 'list_documents');
const spreadsheetIdArgToolNames = new Set([
    'get_spreadsheet',
    'list_sheet_tabs',
    'add_sheet_tab',
    'delete_sheet_tab',
    'read_sheet_values',
    'update_sheet_values',
    'update_timesheet_hours',
    'append_sheet_values',
    'clear_sheet_values'
]);

function isSheetsRelatedToolName(toolName) {
    const name = String(toolName || '');
    if (!name) return false;
    return sheetsToolNames.has(name) || name.startsWith(SHEETS_MCP_TOOL_PREFIX);
}

function applyUserIntentGuardsToToolArgs({ toolName, args, userMessage }) {
    if (!calendarToolNames.has(toolName)) {
        return args;
    }

    const safeArgs = (args && typeof args === 'object' && !Array.isArray(args))
        ? { ...args }
        : {};

    const meetingIntent = hasMeetingIntentInMessage(userMessage);
    const selfOnlyFromMessage = hasSelfOnlyAttendeeIntentInMessage(userMessage);
    const selfOnlyFromArgs = Array.isArray(safeArgs.attendees)
        && safeArgs.attendees.some(item => isSelfOnlyAttendeeDirective(item));
    const forceSelfOnlyAttendees = selfOnlyFromMessage || selfOnlyFromArgs;

    if ((toolName === 'create_event' || toolName === 'create_meet_event') && forceSelfOnlyAttendees) {
        safeArgs.attendees = ['me'];
    }

    if (toolName === 'update_event_attendees' && forceSelfOnlyAttendees) {
        safeArgs.addAttendees = ['me'];
        delete safeArgs.attendees;
    }

    if (toolName === 'create_event' && meetingIntent && safeArgs.createMeetLink !== false) {
        safeArgs.createMeetLink = true;
    }

    return safeArgs;
}

function shouldPreferDocumentEditingRoute(userMessage) {
    const message = String(userMessage || '').toLowerCase();
    if (!message) return false;

    const hasDocKeyword = /\b(doc|docs|document|google doc|docx)\b/i.test(message);
    const hasEditVerb = /\b(edit|update|append|add|insert|replace|modify|change|write)\b/i.test(message);
    const hasSheetKeyword = /\b(sheet|spreadsheet|excel|xlsx|csv|timesheet)\b/i.test(message);

    return hasDocKeyword && hasEditVerb && !hasSheetKeyword;
}

function selectToolsForMessageIntent({ userMessage, availableTools, docsConnected, docsListOnlyConnected }) {
    let selectedTools = Array.isArray(availableTools) ? availableTools : [];
    const routingHints = [];

    if (shouldPreferDocumentEditingRoute(userMessage)) {
        selectedTools = selectedTools.filter(tool => {
            const toolName = tool?.function?.name;
            if (!toolName) return false;
            if (isSheetsRelatedToolName(toolName)) return false;
            if (githubToolNames.has(toolName)) return false;
            return true;
        });
        routingHints.push('Document editing intent detected: avoid Sheets and GitHub file tools.');
        if (!docsConnected && docsListOnlyConnected) {
            routingHints.push('Docs write scope missing: use Drive document append/update fallback tools.');
        }
    }

    if (!selectedTools.length && Array.isArray(availableTools) && availableTools.length > 0) {
        selectedTools = availableTools;
    }

    return { selectedTools, routingHints };
}

function sheetValuesToPlainText(values) {
    if (!Array.isArray(values)) return '';
    const lines = values.map(row => {
        if (Array.isArray(row)) {
            return row.map(cell => String(cell ?? '')).join('\t');
        }
        return String(row ?? '');
    });
    return lines.join('\n').trim();
}

async function resolveDriveFileMetadata(fileId) {
    if (!driveClient || !fileId) return null;
    try {
        const response = await driveClient.files.get({
            fileId: String(fileId),
            supportsAllDrives: true,
            fields: 'id,name,mimeType'
        });
        return response?.data || null;
    } catch {
        return null;
    }
}

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
        create_meet_event: createMeetEvent, add_meet_link_to_event: addMeetLinkToEvent,
        update_event: updateEvent, delete_event: deleteEvent, list_calendars: listCalendars,
        create_calendar: createCalendar, quick_add_event: quickAddEvent,
        get_free_busy: getFreeBusy, check_person_availability: checkPersonAvailability,
        find_common_free_slots: findCommonFreeSlots, list_recurring_instances: listRecurringInstances,
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

async function executeGchatTool(toolName, args) {
    if (!gchatClient) throw new Error('Google Chat not connected. Please authenticate with Google first.');
    const toolMap = {
        list_chat_spaces: listChatSpaces,
        send_chat_message: sendChatMessage,
        list_chat_messages: listChatMessages
    };
    const fn = toolMap[toolName];
    if (!fn) throw new Error(`Unknown Google Chat tool: ${toolName}`);
    try {
        return await fn(args);
    } catch (error) {
        const permissionError = getGchatPermissionError(error);
        if (permissionError) {
            gchatClient = null;
            throw new Error(permissionError);
        }
        throw error;
    }
}

async function executeDriveTool(toolName, args) {
    if (!driveClient) throw new Error('Google Drive not connected. Please authenticate with Google first.');
    const toolMap = {
        list_drive_files: listDriveFiles,
        get_drive_file: getDriveFile,
        create_drive_folder: createDriveFolder,
        create_drive_file: createDriveFile,
        update_drive_file: updateDriveFile,
        delete_drive_file: deleteDriveFile,
        copy_drive_file: copyDriveFile,
        move_drive_file: moveDriveFile,
        share_drive_file: shareDriveFile,
        download_drive_file: downloadDriveFile, download_drive_file_to_local: downloadDriveFileToLocal,
        extract_drive_file_text: extractDriveFileText,
        append_drive_document_text: appendDriveDocumentText,
        convert_file_to_google_doc: convertFileToGoogleDoc,
        convert_file_to_google_sheet: convertFileToGoogleSheet
    };
    const fn = toolMap[toolName];
    if (!fn) throw new Error(`Unknown Google Drive tool: ${toolName}`);
    try {
        return await fn(args);
    } catch (error) {
        const permissionError = getDrivePermissionError(error);
        if (permissionError) {
            driveClient = null;
            throw new Error(permissionError);
        }
        throw error;
    }
}

async function executeSheetsTool(toolName, args) {
    if (!sheetsClient) throw new Error('Google Sheets not connected. Please authenticate with Google first.');
    const toolMap = {
        list_spreadsheets: listSpreadsheets,
        create_spreadsheet: createSpreadsheet,
        get_spreadsheet: getSpreadsheet,
        list_sheet_tabs: listSheetTabs,
        add_sheet_tab: addSheetTab,
        delete_sheet_tab: deleteSheetTab,
        read_sheet_values: readSheetValues,
        update_sheet_values: updateSheetValues,
        update_timesheet_hours: updateTimesheetHours,
        append_sheet_values: appendSheetValues,
        clear_sheet_values: clearSheetValues
    };

    const spreadsheetId = spreadsheetIdArgToolNames.has(toolName)
        ? String(args?.spreadsheetId || '').trim()
        : '';
    if (spreadsheetId) {
        const driveMeta = await resolveDriveFileMetadata(spreadsheetId);
        const driveMimeType = String(driveMeta?.mimeType || '');
        if (driveMimeType && driveMimeType !== 'application/vnd.google-apps.spreadsheet') {
            if (toolName === 'append_sheet_values' && driveMimeType === 'application/vnd.google-apps.document') {
                const appendTextValue = sheetValuesToPlainText(args?.values);
                if (appendTextValue) {
                    return await appendDriveDocumentText({
                        fileId: spreadsheetId,
                        text: appendTextValue
                    });
                }
            }
            throw new Error(`"${driveMeta?.name || spreadsheetId}" is a document (${driveMimeType}), not a spreadsheet. Use Docs/Drive document editing tools.`);
        }
    }

    const fn = toolMap[toolName];
    if (!fn) throw new Error(`Unknown Google Sheets tool: ${toolName}`);
    try {
        return await fn(args);
    } catch (error) {
        const permissionError = getSheetsPermissionError(error);
        if (permissionError) {
            sheetsClient = null;
            throw new Error(permissionError);
        }
        throw error;
    }
}

async function executeGitHubTool(toolName, args) {
    if (!octokitClient) throw new Error('GitHub not connected. Please connect GitHub first.');
    const toolMap = {
        list_repos: listRepos, get_repo: getRepo, create_repo: createRepo,
        list_issues: listIssues, create_issue: createIssue, update_issue: updateIssue,
        list_pull_requests: listPullRequests, get_pull_request: getPullRequest,
        create_pull_request: createPullRequest, merge_pull_request: mergePullRequest,
        list_branches: listBranches, create_branch: createBranch,
        get_file_content: getFileContent, create_or_update_file: createOrUpdateFile,
        search_repos: searchRepos, search_code: searchCode, list_commits: listCommits,
        revert_commit: revertCommit, reset_branch: resetBranch,
        get_user_profile: getUserProfile, list_notifications: listNotifications,
        list_gists: listGists
    };
    const fn = toolMap[toolName];
    if (!fn) throw new Error(`Unknown GitHub tool: ${toolName}`);
    return await fn(args);
}

async function executeOutlookTool(toolName, args) {
    if (!outlookAccessToken) throw new Error('Outlook not connected. Please authenticate first.');
    const toolMap = {
        outlook_send_email: outlookSendEmail, outlook_list_emails: outlookListEmails,
        outlook_read_email: outlookReadEmail, outlook_search_emails: outlookSearchEmails,
        outlook_reply_to_email: outlookReplyToEmail, outlook_forward_email: outlookForwardEmail,
        outlook_delete_email: outlookDeleteEmail, outlook_move_email: outlookMoveEmail,
        outlook_mark_as_read: outlookMarkAsRead, outlook_mark_as_unread: outlookMarkAsUnread,
        outlook_list_folders: outlookListFolders, outlook_create_folder: outlookCreateFolder,
        outlook_get_attachments: outlookGetAttachments, outlook_create_draft: outlookCreateDraft,
        outlook_send_draft: outlookSendDraft, outlook_list_drafts: outlookListDrafts,
        outlook_flag_email: outlookFlagEmail, outlook_get_user_profile: outlookGetUserProfile
    };
    const fn = toolMap[toolName];
    if (!fn) throw new Error(`Unknown Outlook tool: ${toolName}`);
    try {
        return await fn(args);
    } catch (error) {
        const status = error?.status || error?.code;
        const message = String(error?.message || '');
        if (status === 401 || /unauthorized|invalid.*token|expired/i.test(message)) {
            clearOutlookAuth();
            throw new Error('Outlook authentication expired. Please reconnect Outlook.');
        }
        throw error;
    }
}

async function executeDocsTool(toolName, args) {
    if (!docsClient && toolName !== 'list_documents') throw new Error('Google Docs not connected. Please authenticate with Google first.');
    if (toolName === 'list_documents' && !driveClient) throw new Error('Google Drive not connected (required for listing Docs).');
    const toolMap = {
        list_documents: listDocuments, get_document: getDocument, create_document: createDocument,
        insert_text: insertText, replace_text: replaceText, delete_content: deleteContent,
        append_text: appendText, get_document_text: getDocumentText
    };
    const fn = toolMap[toolName];
    if (!fn) throw new Error(`Unknown Google Docs tool: ${toolName}`);
    return await fn(args);
}

async function executeMeetingTranscriptionTool(toolName, args) {
    if (!driveClient) throw new Error('Google Drive not connected. Please authenticate with Google first.');
    const toolMap = {
        list_meeting_transcriptions: listMeetingTranscriptions,
        open_meeting_transcription_file: openMeetingTranscriptionFile,
        summarize_meeting_transcription: summarizeMeetingTranscription
    };
    const fn = toolMap[toolName];
    if (!fn) throw new Error(`Unknown Meeting Transcription tool: ${toolName}`);
    try {
        return await fn(args);
    } catch (error) {
        const permissionError = getDrivePermissionError(error);
        if (permissionError) {
            driveClient = null;
            throw new Error(permissionError);
        }
        throw error;
    }
}

async function executeTeamsTool(toolName, args) {
    if (!outlookAccessToken) throw new Error('Microsoft Teams not connected. Please authenticate with Teams first.');
    const toolMap = {
        teams_list_teams: teamsListTeams, teams_get_team: teamsGetTeam,
        teams_list_channels: teamsListChannels, teams_send_channel_message: teamsSendChannelMessage,
        teams_list_channel_messages: teamsListChannelMessages, teams_list_chats: teamsListChats,
        teams_send_chat_message: teamsSendChatMessage, teams_list_chat_messages: teamsListChatMessages,
        teams_create_chat: teamsCreateChat, teams_get_chat_members: teamsGetChatMembers
    };
    const fn = toolMap[toolName];
    if (!fn) throw new Error(`Unknown Teams tool: ${toolName}`);
    try {
        return await fn(args);
    } catch (error) {
        const status = error?.status || error?.code;
        const message = String(error?.message || '');
        if (status === 401 || /unauthorized|invalid.*token|expired/i.test(message)) {
            throw new Error('Teams authentication expired. Please reconnect Microsoft Teams.');
        }
        throw error;
    }
}

async function executeGcsTool(toolName, args) {
    if (!gcsAuthenticated) throw new Error('GCS not connected. Set GCP_SERVICE_ACCOUNT_KEY_FILE and GCP_PROJECT_ID in .env and restart.');
    const toolMap = {
        gcs_list_buckets: gcsListBuckets, gcs_get_bucket: gcsGetBucket,
        gcs_create_bucket: gcsCreateBucket, gcs_delete_bucket: gcsDeleteBucket,
        gcs_list_objects: gcsListObjects, gcs_upload_object: gcsUploadObject,
        gcs_download_object: gcsDownloadObject, gcs_delete_object: gcsDeleteObject,
        gcs_copy_object: gcsCopyObject, gcs_get_object_metadata: gcsGetObjectMetadata,
        gcs_move_object: gcsMoveObject, gcs_rename_object: gcsRenameObject,
        gcs_make_object_public: gcsMakeObjectPublic, gcs_make_object_private: gcsMakeObjectPrivate,
        gcs_generate_signed_url: gcsGenerateSignedUrl, gcs_batch_delete_objects: gcsBatchDeleteObjects
    };
    const fn = toolMap[toolName];
    if (!fn) throw new Error(`Unknown GCS tool: ${toolName}`);
    return await fn(args);
}

// Master dispatcher
async function executeTool(toolName, args) {
    console.log(`[Tool] ${toolName}`, JSON.stringify(args).slice(0, 200));

    let normalizedArgs = args;
    if (EMAIL_SEND_CONFIRMATION_TOOLS.has(toolName)) {
        normalizedArgs = await prepareEmailSendArgs(toolName, args || {});
    }

    if (EMAIL_SEND_CONFIRMATION_TOOLS.has(toolName) && !hasEmailSendConfirmation(normalizedArgs)) {
        throw new Error(buildEmailSendConfirmationMessage(toolName, normalizedArgs || {}));
    }

    const sanitizedArgs = EMAIL_SEND_CONFIRMATION_TOOLS.has(toolName)
        ? stripSendConfirmationFlags(normalizedArgs)
        : normalizedArgs;

    if (gmailToolNames.has(toolName)) return await executeGmailTool(toolName, sanitizedArgs);
    if (calendarToolNames.has(toolName)) return await executeCalendarTool(toolName, sanitizedArgs);
    if (gchatToolNames.has(toolName)) return await executeGchatTool(toolName, sanitizedArgs);
    if (driveToolNames.has(toolName)) return await executeDriveTool(toolName, sanitizedArgs);
    if (sheetsToolNames.has(toolName)) return await executeSheetsTool(toolName, sanitizedArgs);
    if (toolName.startsWith(SHEETS_MCP_TOOL_PREFIX)) return await executeSheetsMcpTool(toolName, sanitizedArgs);
    if (docsToolNames.has(toolName)) return await executeDocsTool(toolName, sanitizedArgs);
    if (meetingTranscriptionToolNames.has(toolName)) return await executeMeetingTranscriptionTool(toolName, sanitizedArgs);
    if (githubToolNames.has(toolName)) return await executeGitHubTool(toolName, sanitizedArgs);
    if (outlookToolNames.has(toolName)) return await executeOutlookTool(toolName, sanitizedArgs);
    if (teamsToolNames.has(toolName)) return await executeTeamsTool(toolName, sanitizedArgs);
    if (gcsToolNames.has(toolName)) return await executeGcsTool(toolName, sanitizedArgs);
    throw new Error(`Unknown tool: ${toolName}`);
}

// ============================================================
//  API ROUTES
// ============================================================

// Return tools list for the UI grouped by service
// Serve downloaded files for browser download
app.get('/api/download/:filename', (req, res) => {
    try {
        const filename = req.params.filename;
        const filepath = path.join(UPLOADS_DIR, filename);

        // Security: use resolve + validateLocalPath to prevent path traversal
        if (!validateLocalPath(filepath)) {
            return res.status(403).json({ error: 'Access denied' });
        }

        // Check if file exists
        if (!fs.existsSync(filepath)) {
            return res.status(404).json({ error: 'File not found' });
        }

        // Extract original filename (remove timestamp and hash)
        const originalFilename = filename.replace(/-\d+-[a-f0-9]+\./, '.');

        // Send file with download headers
        res.download(filepath, originalFilename, (err) => {
            if (err) {
                console.error('Error sending file:', err);
                if (!res.headersSent) {
                    res.status(500).json({ error: 'Failed to download file' });
                }
            }
        });
    } catch (error) {
        console.error('Download error:', error);
        res.status(500).json({ error: 'Internal server error' });
    }
});

app.get('/api/tools', (req, res) => {
    const gmail = { service: 'gmail', connected: !!gmailClient, tools: gmailTools.map(t => ({ function: t.function })) };
    const calendarConnected = !!calendarClient && hasCalendarScope();
    const calendar = { service: 'calendar', connected: calendarConnected, tools: calendarTools.map(t => ({ function: t.function })) };
    const gchatConnected = !!gchatClient && hasGchatScopes();
    const gchat = { service: 'gchat', connected: gchatConnected, tools: gchatTools.map(t => ({ function: t.function })) };
    const driveConnected = !!driveClient && hasDriveScope();
    const drive = { service: 'drive', connected: driveConnected, tools: driveTools.map(t => ({ function: t.function })) };
    const sheetsConnected = !!sheetsClient && hasSheetsScope();
    const sheets = {
        service: 'sheets',
        connected: sheetsConnected,
        tools: [...sheetsTools, ...sheetsMcpTools].map(t => ({ function: t.function }))
    };
    const github = { service: 'github', connected: !!octokitClient, tools: githubTools.map(t => ({ function: t.function })) };
    const outlook = { service: 'outlook', connected: !!outlookAccessToken, tools: outlookTools.map(t => ({ function: t.function })) };
    const docsConnected = !!docsClient && hasDocsScope();
    const docsListOnlyConnected = !docsConnected && driveConnected;
    const docsVisibleTools = docsConnected ? docsTools : (docsListOnlyConnected ? docsListOnlyTools : []);
    const docs = { service: 'docs', connected: docsConnected || docsListOnlyConnected, tools: docsVisibleTools.map(t => ({ function: t.function })) };
    const meetingTranscription = {
        service: 'meeting_transcription',
        connected: driveConnected,
        tools: meetingTranscriptionTools.map(t => ({ function: t.function }))
    };
    const teamsConnected = !!outlookAccessToken && hasTeamsScopes();
    const teams = { service: 'teams', connected: teamsConnected, tools: teamsTools.map(t => ({ function: t.function })) };
    const gcs = { service: 'gcs', connected: gcsAuthenticated, tools: gcsTools.map(t => ({ function: t.function })) };
    const totalTools = gmailTools.length + calendarTools.length + gchatTools.length + driveTools.length + sheetsTools.length + sheetsMcpTools.length + githubTools.length + outlookTools.length + docsTools.length + meetingTranscriptionTools.length + teamsTools.length + gcsTools.length;
    res.json({
        services: [gmail, calendar, gchat, drive, sheets, github, outlook, docs, meetingTranscription, teams, gcs],
        totalTools
    });
});

app.get('/api/timer-tasks/status', (req, res) => {
    const enabledCount = scheduledTasks.filter(task => task.enabled).length;
    const runningCount = scheduledTasks.filter(task => runningScheduledTaskIds.has(task.id)).length;
    res.json({
        connected: true,
        taskCount: scheduledTasks.length,
        enabledCount,
        runningCount
    });
});

app.get('/api/timer-tasks', (req, res) => {
    const tasks = scheduledTasks
        .map(task => ({
            ...task,
            running: runningScheduledTaskIds.has(task.id)
        }))
        .sort((a, b) => a.time.localeCompare(b.time));
    res.json({ tasks, count: tasks.length });
});

app.post('/api/timer-tasks', (req, res) => {
    try {
        const task = sanitizeScheduledTask({ ...req.body }, { forCreate: true });
        if (scheduledTasks.some(existing => existing.id === task.id)) {
            task.id = makeId('task');
        }
        scheduledTasks.push(task);
        saveScheduledTasksToDisk();
        res.status(201).json({ success: true, task });
    } catch (error) {
        res.status(400).json({ error: 'Invalid request' });
    }
});

app.patch('/api/timer-tasks/:id', (req, res) => {
    try {
        const index = scheduledTasks.findIndex(task => task.id === req.params.id);
        if (index < 0) return res.status(404).json({ error: 'Task not found' });

        const existing = scheduledTasks[index];
        const merged = sanitizeScheduledTask({
            ...existing,
            ...req.body,
            id: existing.id,
            createdAt: existing.createdAt
        });
        scheduledTasks[index] = merged;
        saveScheduledTasksToDisk();
        res.json({ success: true, task: merged });
    } catch (error) {
        res.status(400).json({ error: 'Invalid request' });
    }
});

app.delete('/api/timer-tasks/:id', (req, res) => {
    const before = scheduledTasks.length;
    scheduledTasks = scheduledTasks.filter(task => task.id !== req.params.id);
    if (scheduledTasks.length === before) {
        return res.status(404).json({ error: 'Task not found' });
    }
    saveScheduledTasksToDisk();
    res.json({ success: true });
});

app.post('/api/timer-tasks/:id/run', async (req, res) => {
    try {
        const task = scheduledTasks.find(item => item.id === req.params.id);
        if (!task) return res.status(404).json({ error: 'Task not found' });
        const result = await runScheduledTask(task.id, 'manual');
        if (!result.success && !result.skipped) {
            return res.status(500).json(result);
        }
        return res.json(result);
    } catch (error) {
        return res.status(500).json({ error: 'Failed to run timer task' });
    }
});

// Gmail authentication status
app.get('/api/gmail/status', (req, res) => {
    const hasCredentials = isGoogleOAuthConfigured();
    const hasToken = fs.existsSync(TOKEN_PATH);
    res.json({
        credentialsConfigured: !!hasCredentials,
        authenticated: hasToken && gmailClient !== null,
        toolCount: gmailTools.length
    });
});

// Calendar authentication status (same as Gmail since same OAuth)
app.get('/api/calendar/status', (req, res) => {
    const hasCredentials = isGoogleOAuthConfigured();
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

// Google Chat authentication status (same Google OAuth, separate scopes)
app.get('/api/gchat/status', (req, res) => {
    const hasCredentials = isGoogleOAuthConfigured();
    const hasToken = fs.existsSync(TOKEN_PATH);
    const gchatScopeGranted = hasGchatScopes();
    res.json({
        credentialsConfigured: !!hasCredentials,
        authenticated: hasToken && gchatClient !== null && gchatScopeGranted,
        hasGchatScopes: gchatScopeGranted,
        requiresReconnect: hasToken && !gchatScopeGranted,
        toolCount: gchatTools.length
    });
});

// Google Drive authentication status (same Google OAuth, separate scope)
app.get('/api/drive/status', (req, res) => {
    const hasCredentials = isGoogleOAuthConfigured();
    const hasToken = fs.existsSync(TOKEN_PATH);
    const driveScopeGranted = hasDriveScope();
    res.json({
        credentialsConfigured: !!hasCredentials,
        authenticated: hasToken && driveClient !== null && driveScopeGranted,
        hasDriveScope: driveScopeGranted,
        requiresReconnect: hasToken && !driveScopeGranted,
        toolCount: driveTools.length
    });
});

// Stream/download a Drive file (raw bytes for normal files, export for Google Workspace files)
app.get('/api/drive/download/:fileId', async (req, res) => {
    try {
        if (!driveClient) {
            return res.status(401).json({ error: 'Google Drive not connected' });
        }
        const fileId = String(req.params.fileId || '').trim();
        if (!fileId) return res.status(400).json({ error: 'fileId is required' });

        const requestedFormat = String(req.query.format || '').trim().toLowerCase();
        const requestedFilename = String(req.query.filename || '').trim();

        const meta = await driveClient.files.get({
            fileId,
            supportsAllDrives: true,
            fields: 'id,name,mimeType,size'
        });
        const sourceMimeType = String(meta.data.mimeType || 'application/octet-stream');
        let contentType = sourceMimeType;
        let payload;
        let downloadName = sanitizeDownloadFilename(requestedFilename || meta.data.name || 'download');

        if (sourceMimeType.startsWith('application/vnd.google-apps.')) {
            const exportInfo = resolveDriveExportFormat(sourceMimeType, requestedFormat);
            if (!exportInfo) {
                return res.status(400).json({
                    error: `File type "${sourceMimeType}" does not support requested export format "${requestedFormat || 'default'}".`
                });
            }
            const exportResponse = await driveClient.files.export(
                { fileId, mimeType: exportInfo.mimeType },
                { responseType: 'arraybuffer' }
            );
            payload = Buffer.from(exportResponse.data);
            contentType = exportInfo.mimeType;
            downloadName = ensureFilenameExtension(downloadName, exportInfo.extension);
        } else {
            const fileResponse = await driveClient.files.get(
                { fileId, alt: 'media', supportsAllDrives: true },
                { responseType: 'arraybuffer' }
            );
            payload = Buffer.from(fileResponse.data);
        }

        const asciiFileName = downloadName
            .replace(/[\r\n]/g, '')
            .replace(/[^\x20-\x7E]/g, '_')
            .replace(/["\\]/g, '_');
        res.setHeader('Content-Type', contentType || 'application/octet-stream');
        res.setHeader('Content-Length', String(payload.length));
        res.setHeader('Cache-Control', 'no-store');
        res.setHeader('Content-Disposition', `attachment; filename="${asciiFileName}"; filename*=UTF-8''${encodeURIComponent(downloadName)}`);
        return res.status(200).send(payload);
    } catch (error) {
        console.error('Drive download error:', error?.message || error);
        const permissionError = getDrivePermissionError(error);
        if (permissionError) {
            driveClient = null;
            return res.status(401).json({ error: permissionError });
        }
        return res.status(500).json({ error: 'Failed to download Drive file' });
    }
});

// GCS download endpoint
app.get('/api/gcs/download/:bucket/:objectName(*)', async (req, res) => {
    try {
        if (!gcsAuthenticated || !gcsClient) {
            return res.status(401).json({ error: 'GCS not connected' });
        }
        const bucket = String(req.params.bucket || '').trim();
        const objectName = String(req.params.objectName || '').trim();
        if (!bucket || !objectName) {
            return res.status(400).json({ error: 'bucket and objectName are required' });
        }

        const requestedFilename = String(req.query.filename || '').trim();

        // Get object metadata
        const meta = await gcsClient.objects.get({ bucket, object: objectName });
        const contentType = meta.data.contentType || 'application/octet-stream';
        const downloadName = requestedFilename || objectName.split('/').pop() || 'download';

        // Download the file content
        const fileResponse = await gcsClient.objects.get(
            { bucket, object: objectName, alt: 'media' },
            { responseType: 'arraybuffer' }
        );
        const payload = Buffer.from(fileResponse.data);

        // Sanitize filename for Content-Disposition
        const asciiFileName = downloadName
            .replace(/[\r\n]/g, '')
            .replace(/[^\x20-\x7E]/g, '_')
            .replace(/["\\]/g, '_');

        res.setHeader('Content-Type', contentType);
        res.setHeader('Content-Length', String(payload.length));
        res.setHeader('Cache-Control', 'no-store');
        res.setHeader('Content-Disposition', `attachment; filename="${asciiFileName}"; filename*=UTF-8''${encodeURIComponent(downloadName)}`);
        return res.status(200).send(payload);
    } catch (error) {
        console.error('GCS download error:', error?.message || error);
        return res.status(500).json({ error: 'Failed to download GCS object' });
    }
});

// Google Sheets authentication status (same Google OAuth, separate scope)
app.get('/api/sheets/status', (req, res) => {
    const hasCredentials = isGoogleOAuthConfigured();
    const hasToken = fs.existsSync(TOKEN_PATH);
    const sheetsScopeGranted = hasSheetsScope();
    res.json({
        credentialsConfigured: !!hasCredentials,
        authenticated: hasToken && sheetsClient !== null && sheetsScopeGranted,
        hasSheetsScope: sheetsScopeGranted,
        requiresReconnect: hasToken && !sheetsScopeGranted,
        toolCount: sheetsTools.length
    });
});

app.get('/api/sheets-mcp/status', (req, res) => {
    res.json({
        enabled: SHEETS_MCP_ENABLED,
        connected: !!sheetsMcpClient,
        toolCount: sheetsMcpTools.length,
        error: sheetsMcpError ? 'Connection error' : null
    });
});

app.post('/api/sheets-mcp/reconnect', async (req, res) => {
    try {
        await closeSheetsMcpClient();
        await initSheetsMcpClient();
        res.json({
            success: !!sheetsMcpClient,
            connected: !!sheetsMcpClient,
            toolCount: sheetsMcpTools.length,
            error: sheetsMcpError
        });
    } catch (error) {
        res.status(500).json({
            success: false,
            connected: false,
            toolCount: sheetsMcpTools.length,
            error: error.message || 'Failed to reconnect Sheets MCP'
        });
    }
});

// Google Docs status (same Google OAuth, separate scope)
app.get('/api/docs/status', (req, res) => {
    const hasCredentials = isGoogleOAuthConfigured();
    const hasToken = fs.existsSync(TOKEN_PATH);
    const docsScopeGranted = hasDocsScope();
    res.json({
        credentialsConfigured: !!hasCredentials,
        authenticated: hasToken && docsClient !== null && docsScopeGranted,
        hasDocsScope: docsScopeGranted,
        requiresReconnect: hasToken && !docsScopeGranted,
        toolCount: docsTools.length
    });
});

app.get('/api/meeting-transcription/status', (req, res) => {
    const hasCredentials = isGoogleOAuthConfigured();
    const hasToken = fs.existsSync(TOKEN_PATH);
    const driveScopeGranted = hasDriveScope();
    res.json({
        credentialsConfigured: !!hasCredentials,
        authenticated: hasToken && driveClient !== null && driveScopeGranted,
        hasDriveScope: driveScopeGranted,
        requiresReconnect: hasToken && !driveScopeGranted,
        toolCount: meetingTranscriptionTools.length
    });
});

// Microsoft Teams status (uses Outlook OAuth with extra scopes)
app.get('/api/teams/status', (req, res) => {
    const oauthConfigured = isOutlookOAuthConfigured();
    const teamsScoped = hasTeamsScopes();
    res.json({
        authenticated: !!outlookAccessToken && teamsScoped,
        oauthConfigured,
        hasScopes: teamsScoped,
        email: outlookUserEmail,
        requiresReconnect: !!outlookAccessToken && !teamsScoped,
        toolCount: teamsTools.length
    });
});

// GCS (Cloud Storage) status
app.get('/api/gcs/status', (req, res) => {
    res.json({
        authenticated: gcsAuthenticated,
        projectId: gcsProjectId,
        toolCount: gcsTools.length
    });
});

// Microsoft Teams auth - triggers Outlook re-auth with Teams scopes
app.get('/api/teams/auth', (req, res) => {
    if (!isOutlookOAuthConfigured()) {
        return res.status(400).json({
            error: 'Outlook OAuth credentials are not configured in .env file (required for Teams)',
            setupRequired: true
        });
    }
    const state = issueOutlookOAuthState();
    const params = new URLSearchParams({
        client_id: process.env.OUTLOOK_CLIENT_ID,
        response_type: 'code',
        redirect_uri: getOutlookRedirectUri(),
        scope: OUTLOOK_SCOPES.join(' '),
        response_mode: 'query',
        state
    });
    const authUrl = `${OUTLOOK_AUTHORITY}/oauth2/v2.0/authorize?${params.toString()}`;
    res.json({ authUrl });
});

// Calendar connect - triggers re-auth with calendar scope
app.get('/api/calendar/connect', (req, res) => {
    if (!oauth2Client) {
        const initialized = initOAuthClient();
        if (!initialized) {
            return res.status(400).json({ error: 'Google OAuth credentials not configured in .env file', setupRequired: true });
        }
    }
    const state = issueGoogleOAuthState();
    const authUrl = oauth2Client.generateAuthUrl({ access_type: 'offline', scope: SCOPES, prompt: GOOGLE_OAUTH_PROMPT, state });
    res.json({ authUrl });
});

// Google Chat connect - triggers re-auth with Chat scopes
app.get('/api/gchat/connect', (req, res) => {
    if (!oauth2Client) {
        const initialized = initOAuthClient();
        if (!initialized) {
            return res.status(400).json({ error: 'Google OAuth credentials not configured in .env file', setupRequired: true });
        }
    }
    const state = issueGoogleOAuthState();
    const authUrl = oauth2Client.generateAuthUrl({ access_type: 'offline', scope: SCOPES, prompt: GOOGLE_OAUTH_PROMPT, state });
    res.json({ authUrl });
});

// Google Drive connect - triggers re-auth with Drive scope
app.get('/api/drive/connect', (req, res) => {
    if (!oauth2Client) {
        const initialized = initOAuthClient();
        if (!initialized) {
            return res.status(400).json({ error: 'Google OAuth credentials not configured in .env file', setupRequired: true });
        }
    }
    const state = issueGoogleOAuthState();
    const authUrl = oauth2Client.generateAuthUrl({ access_type: 'offline', scope: SCOPES, prompt: GOOGLE_OAUTH_PROMPT, state });
    res.json({ authUrl });
});

// Google Sheets connect - triggers re-auth with Sheets scope
app.get('/api/sheets/connect', (req, res) => {
    if (!oauth2Client) {
        const initialized = initOAuthClient();
        if (!initialized) {
            return res.status(400).json({ error: 'Google OAuth credentials not configured in .env file', setupRequired: true });
        }
    }
    const state = issueGoogleOAuthState();
    const authUrl = oauth2Client.generateAuthUrl({ access_type: 'offline', scope: SCOPES, prompt: GOOGLE_OAUTH_PROMPT, state });
    res.json({ authUrl });
});

// GitHub authentication status
app.get('/api/github/status', (req, res) => {
    res.json({
        authenticated: octokitClient !== null,
        oauthConfigured: isGitHubOAuthConfigured(),
        username: githubUsername,
        authMethod: githubAuthMethod,
        toolCount: githubTools.length
    });
});

// GitHub OAuth connect
app.get('/api/github/auth', (req, res) => {
    if (!isGitHubOAuthConfigured()) {
        return res.status(400).json({
            error: 'GitHub OAuth credentials are not configured in .env file',
            setupRequired: true
        });
    }
    const state = issueGithubOAuthState();
    const params = new URLSearchParams({
        client_id: process.env.GITHUB_CLIENT_ID,
        redirect_uri: getGitHubRedirectUri(),
        scope: GITHUB_OAUTH_SCOPES.join(' '),
        state
    });
    const authUrl = `https://github.com/login/oauth/authorize?${params.toString()}`;
    res.json({ authUrl });
});

// GitHub connect with PAT
app.post('/api/github/connect', async (req, res) => {
    const { token } = req.body;
    if (!token) return res.status(400).json({ error: 'Token is required' });

    try {
        const testClient = new Octokit({ auth: token });
        const user = await testClient.rest.users.getAuthenticated();
        octokitClient = testClient;
        githubUsername = user.data.login;
        githubAuthMethod = 'pat';

        // Save token
        saveGitHubTokenData({
            token,
            username: githubUsername,
            authMethod: githubAuthMethod,
            connectedAt: new Date().toISOString()
        });

        res.json({ success: true, username: user.data.login, message: `Connected as ${user.data.login}` });
    } catch (error) {
        res.status(401).json({ error: 'Invalid or expired token' });
    }
});

// GitHub disconnect
app.post('/api/github/disconnect', (req, res) => {
    clearGitHubAuth();
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
    const state = issueGoogleOAuthState();
    const authUrl = oauth2Client.generateAuthUrl({ access_type: 'offline', scope: SCOPES, prompt: GOOGLE_OAUTH_PROMPT, state });
    res.json({ authUrl });
});

// GitHub OAuth callback
app.get('/github/callback', async (req, res) => {
    const { code, state, error, error_description: errorDescription } = req.query;

    if (error) {
        return res.status(400).send(`GitHub authentication failed: ${errorDescription || error}`);
    }
    if (!code || !state) {
        return res.status(400).send('Missing GitHub authorization code or state');
    }
    if (!consumeGithubOAuthState(state)) {
        return res.status(400).send('Invalid or expired GitHub OAuth state. Please try connecting again.');
    }
    if (!isGitHubOAuthConfigured()) {
        return res.status(400).send('GitHub OAuth credentials are not configured in server .env');
    }

    try {
        const tokenResponse = await fetch('https://github.com/login/oauth/access_token', {
            method: 'POST',
            headers: {
                'Accept': 'application/json',
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                client_id: process.env.GITHUB_CLIENT_ID,
                client_secret: process.env.GITHUB_CLIENT_SECRET,
                code,
                redirect_uri: getGitHubRedirectUri(),
                state
            })
        });
        const tokenData = await tokenResponse.json();
        if (!tokenResponse.ok || tokenData.error || !tokenData.access_token) {
            const details = tokenData.error_description || tokenData.error || tokenResponse.statusText;
            console.error('GitHub token exchange failed:', details);
            return res.status(401).send('GitHub authentication failed. Please try again.');
        }

        const token = tokenData.access_token;
        const client = new Octokit({ auth: token });
        const user = await client.rest.users.getAuthenticated();

        octokitClient = client;
        githubUsername = user.data.login;
        githubAuthMethod = 'oauth';
        saveGitHubTokenData({
            token,
            username: githubUsername,
            authMethod: githubAuthMethod,
            scope: tokenData.scope || '',
            connectedAt: new Date().toISOString()
        });

        res.send('<html><body style="background:#0f0f1a;color:#fff;display:flex;align-items:center;justify-content:center;height:100vh;font-family:sans-serif"><div style="text-align:center"><h1>GitHub Connected!</h1><p>GitHub tools are ready. You can close this window.</p></div></body></html>');
    } catch (oauthError) {
        console.error('GitHub OAuth callback error:', oauthError);
        res.status(500).send(`GitHub authentication failed: ${oauthError.message}`);
    }
});

// OAuth callback
app.get('/oauth2callback', async (req, res) => {
    const { code, state } = req.query;
    if (!code || !state) return res.status(400).send('Missing Google authorization code or state');
    if (!consumeGoogleOAuthState(state)) {
        return res.status(400).send('Invalid or expired Google OAuth state. Please try connecting again.');
    }
    if (!oauth2Client) {
        const initialized = initOAuthClient();
        if (!initialized) {
            return res.status(400).send('Google OAuth credentials are not configured in server .env');
        }
    }

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
        gchatClient = tokenHasScopes(tokens, GCHAT_REQUIRED_SCOPES)
            ? google.chat({ version: 'v1', auth: oauth2Client })
            : null;
        driveClient = tokenHasScope(tokens, DRIVE_SCOPE)
            ? google.drive({ version: 'v3', auth: oauth2Client })
            : null;
        sheetsClient = tokenHasScope(tokens, SHEETS_SCOPE)
            ? google.sheets({ version: 'v4', auth: oauth2Client })
            : null;
        docsClient = tokenHasScope(tokens, DOCS_SCOPE)
            ? google.docs({ version: 'v1', auth: oauth2Client })
            : null;
        const calendarMessage = calendarClient
            ? 'Gmail + Calendar are ready.'
            : 'Gmail is ready. Calendar permission is still missing, so reconnect Calendar from the app.';
        const gchatMessage = gchatClient
            ? 'Google Chat is ready.'
            : 'Google Chat permission is still missing, so reconnect Chat from the app.';
        const driveMessage = driveClient
            ? 'Google Drive is ready.'
            : 'Google Drive permission is still missing, so reconnect Drive from the app.';
        const sheetsMessage = sheetsClient
            ? 'Google Sheets is ready.'
            : 'Google Sheets permission is still missing, so reconnect Sheets from the app.';
        const docsMessage = docsClient
            ? 'Google Docs is ready.'
            : 'Google Docs permission is still missing, so reconnect Docs from the app.';
        res.send(`<html><body style="background:#0f0f1a;color:#fff;display:flex;align-items:center;justify-content:center;height:100vh;font-family:sans-serif"><div style="text-align:center"><h1>Google Connected!</h1><p>${calendarMessage} ${gchatMessage} ${driveMessage} ${sheetsMessage} ${docsMessage} You can close this window.</p></div></body></html>`);
    } catch (error) {
        console.error('OAuth callback error:', error);
        res.status(500).send(`Authentication failed: ${error.message}`);
    }
});

// Outlook authentication status
app.get('/api/outlook/status', (req, res) => {
    res.json({
        authenticated: !!outlookAccessToken,
        oauthConfigured: isOutlookOAuthConfigured(),
        email: outlookUserEmail,
        authMethod: outlookAuthMethod,
        toolCount: outlookTools.length
    });
});

// Outlook OAuth connect
app.get('/api/outlook/auth', (req, res) => {
    if (!isOutlookOAuthConfigured()) {
        return res.status(400).json({
            error: 'Outlook OAuth credentials are not configured in .env file',
            setupRequired: true
        });
    }
    const state = issueOutlookOAuthState();
    const params = new URLSearchParams({
        client_id: process.env.OUTLOOK_CLIENT_ID,
        response_type: 'code',
        redirect_uri: getOutlookRedirectUri(),
        scope: OUTLOOK_SCOPES.join(' '),
        response_mode: 'query',
        state
    });
    const authUrl = `${OUTLOOK_AUTHORITY}/oauth2/v2.0/authorize?${params.toString()}`;
    res.json({ authUrl });
});

// Outlook disconnect
app.post('/api/outlook/disconnect', (req, res) => {
    clearOutlookAuth();
    res.json({ success: true, message: 'Outlook disconnected' });
});

// Outlook OAuth callback
app.get('/outlook/callback', async (req, res) => {
    const { code, state, error, error_description: errorDescription } = req.query;

    if (error) {
        return res.status(400).send(`Outlook authentication failed: ${errorDescription || error}`);
    }
    if (!code || !state) {
        return res.status(400).send('Missing Outlook authorization code or state');
    }
    if (!consumeOutlookOAuthState(state)) {
        return res.status(400).send('Invalid or expired Outlook OAuth state. Please try connecting again.');
    }
    if (!isOutlookOAuthConfigured()) {
        return res.status(400).send('Outlook OAuth credentials are not configured in server .env');
    }

    try {
        const tokenParams = new URLSearchParams({
            client_id: process.env.OUTLOOK_CLIENT_ID,
            client_secret: process.env.OUTLOOK_CLIENT_SECRET,
            code,
            redirect_uri: getOutlookRedirectUri(),
            grant_type: 'authorization_code',
            scope: OUTLOOK_SCOPES.join(' ')
        });
        const tokenResponse = await fetch(`${OUTLOOK_AUTHORITY}/oauth2/v2.0/token`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
            body: tokenParams.toString()
        });
        const tokenData = await tokenResponse.json();

        if (!tokenResponse.ok || tokenData.error || !tokenData.access_token) {
            const details = tokenData.error_description || tokenData.error || tokenResponse.statusText;
            console.error('Outlook token exchange failed:', details);
            return res.status(401).send('Outlook authentication failed. Please try again.');
        }

        outlookAccessToken = tokenData.access_token;
        outlookRefreshToken = tokenData.refresh_token || null;
        outlookTokenExpiry = Date.now() + (tokenData.expires_in * 1000) - 60000;

        // Fetch user profile
        const userResp = await fetch(`${OUTLOOK_GRAPH_BASE}/me?$select=displayName,mail,userPrincipalName`, {
            headers: { 'Authorization': `Bearer ${outlookAccessToken}` }
        });
        const userData = userResp.ok ? await userResp.json() : {};
        outlookUserEmail = userData.mail || userData.userPrincipalName || null;

        saveOutlookTokenData({
            access_token: outlookAccessToken,
            refresh_token: outlookRefreshToken,
            expiry: outlookTokenExpiry,
            email: outlookUserEmail,
            displayName: userData.displayName || '',
            scope: tokenData.scope || OUTLOOK_SCOPES.join(' '),
            connectedAt: new Date().toISOString()
        });

        res.send('<html><body style="background:#0f0f1a;color:#fff;display:flex;align-items:center;justify-content:center;height:100vh;font-family:sans-serif"><div style="text-align:center"><h1>Outlook Connected!</h1><p>Outlook tools are ready. You can close this window.</p></div></body></html>');
    } catch (oauthError) {
        console.error('Outlook OAuth callback error:', oauthError);
        res.status(500).send(`Outlook authentication failed: ${oauthError.message}`);
    }
});

// ============================================================
//  AGENTIC CHAT ENDPOINT — Robust multi-turn tool loop
// ============================================================
function getConnectedToolContext() {
    const availableTools = [];
    if (gmailClient) availableTools.push(...gmailTools);
    const calendarConnected = !!calendarClient && hasCalendarScope();
    if (calendarConnected) availableTools.push(...calendarTools);
    const gchatConnected = !!gchatClient && hasGchatScopes();
    if (gchatConnected) availableTools.push(...gchatTools);
    const driveConnected = !!driveClient && hasDriveScope();
    if (driveConnected) availableTools.push(...driveTools);
    const sheetsConnected = !!sheetsClient && hasSheetsScope();
    if (sheetsConnected) availableTools.push(...sheetsTools);
    const sheetsMcpConnected = !!sheetsMcpClient && sheetsMcpTools.length > 0;
    if (sheetsMcpConnected) availableTools.push(...sheetsMcpTools);
    const docsConnected = !!docsClient && hasDocsScope();
    const docsListOnlyConnected = !docsConnected && driveConnected;
    if (docsConnected) availableTools.push(...docsTools);
    else if (docsListOnlyConnected) availableTools.push(...docsListOnlyTools);
    if (driveConnected) availableTools.push(...meetingTranscriptionTools);
    if (octokitClient) availableTools.push(...githubTools);
    if (outlookAccessToken) availableTools.push(...outlookTools);
    const teamsConnected = !!outlookAccessToken && hasTeamsScopes();
    if (teamsConnected) availableTools.push(...teamsTools);
    if (gcsAuthenticated) availableTools.push(...gcsTools);

    const connectedServices = [];
    if (gmailClient) connectedServices.push('Gmail (25 tools)');
    if (calendarConnected) connectedServices.push(`Google Calendar (${calendarTools.length} tools)`);
    if (gchatConnected) connectedServices.push(`Google Chat (${gchatTools.length} tools)`);
    if (driveConnected) connectedServices.push(`Google Drive (${driveTools.length} tools)`);
    if (sheetsConnected) connectedServices.push(`Google Sheets (${sheetsTools.length} tools)`);
    if (sheetsMcpConnected) connectedServices.push(`Google Sheets MCP (${sheetsMcpTools.length} tools)`);
    if (docsConnected) connectedServices.push(`Google Docs (${docsTools.length} tools)`);
    if (docsListOnlyConnected) connectedServices.push(`Google Docs (${docsListOnlyTools.length} tool via Drive scope)`);
    if (driveConnected) connectedServices.push(`Meeting Transcription (${meetingTranscriptionTools.length} tools)`);
    if (octokitClient) connectedServices.push('GitHub (20 tools)');
    if (outlookAccessToken) connectedServices.push(`Outlook (${outlookTools.length} tools)`);
    if (teamsConnected) connectedServices.push(`Microsoft Teams (${teamsTools.length} tools)`);
    if (gcsAuthenticated) connectedServices.push(`GCS (${gcsTools.length} tools)`);

    const statusText = connectedServices.length > 0 ? connectedServices.join(', ') : 'No services connected';
    return { availableTools, statusText, docsConnected, docsListOnlyConnected };
}

function buildAgentSystemPrompt({ statusText, toolCount, dateContext, connectedServices }) {
    // Build list of available service names for clarity
    const availableServicesList = connectedServices && connectedServices.length > 0
        ? connectedServices.join(', ')
        : 'No services';

    const basePrompt = `You are a powerful AI assistant integrated with Gmail, Google Calendar, Google Chat, Google Drive, Google Sheets, Google Docs, Meeting Transcription, GitHub, Outlook, Microsoft Teams, and GCP Cloud Storage. You execute complex, multi-step operations across connected services.

Connected Services: ${statusText}
Total Tools Available: ${toolCount}

## EXECUTION APPROACH
Complete each request end-to-end. Execute tools, evaluate results, continue until done. Only ask when a required value is truly missing or an action is destructive.

---

## CORE RULES

**1. DISCOVERY FIRST**
Never invent IDs, email addresses, or repository names. Always use search/list/discovery tools first when the user refers to emails, files, issues, or repos by description.

**2. FILE SEARCH & SELECTION**
- Search Drive (list_drive_files) or GCS (gcs_list_objects) before any file download or attachment.
- If search returns **multiple matches**: stop and ask the user which file they want. List each option with name and type. Never auto-select.
- If search returns **one match**: proceed immediately without asking.
- In your response, do not enumerate search results — just confirm what was found and what was done.

**3. CALENDAR EVENT VALIDATION**
Before calling create_event or create_meet_event:
- If the meeting title is missing or generic (e.g. "meeting", "event") → ask: "What should I name this meeting?"
- If attendees are missing and it is not a personal reminder/block → ask: "Who should I invite?"
- Do not route calendar requests to Google Docs tools.
- Skip attendee check only if the user explicitly says "personal", "just me", "reminder", or "block time".

**4. EMAIL SEND CONFIRMATION**
Before any send action (send_email, reply_to_email, forward_email, send_draft, outlook_send_email, outlook_reply_to_email, outlook_forward_email, outlook_send_draft):
- Show recipient(s), subject, and body preview. Ask: "Shall I send this?"
- On user confirmation, call the tool with confirmSend: true immediately. Do not ask again.
- If the tool returns a "confirmation required" message, present details cleanly — never echo raw tool text.

**5. EMAIL RECIPIENTS**
When the user gives a person's name (not a full email address):
- Pass the raw name directly to send_email/create_draft (e.g. to: ["kamalakar"]). The backend resolves it from Gmail history.
- Never construct or guess an email address (e.g. do not append @domain.com yourself).
- Never call search_emails to find a recipient's address — pass the name and let the system resolve it.
- If resolution fails, ask: "I couldn't find an email for [name]. Could you provide their exact address?"

**6. ATTACHING DRIVE / GCS FILES TO EMAIL**
Silently: 1) list_drive_files/gcs_list_objects → 2) download_drive_file_to_local/gcs_download_object → 3) send_email with attachments: [{ localPath }].
- Use the EXACT localPath from step 2. Never skip step 2. Never use download_drive_file (browser URL, not local path).
- On confirmation re-send: reuse existing localPath. Do NOT repeat steps 1-2.

**7. COMPOUND TASKS**
Chain all steps in one flow: e.g. search_emails → read_email → reply_to_email.

**8. PARALLELISE / SEQUENCE**
Run independent subtasks simultaneously. Chain dependent steps in order.

**9. NEVER STOP MID-TASK**
Continue until the goal is complete. No unnecessary confirmations on non-destructive steps.

**10. RECOVER FROM FAILURES**
If a tool fails, run a discovery tool and retry. Surface blockers only after retries are exhausted.

**11. DESTRUCTIVE ACTION SAFETY**
Confirm delete/trash/clear/bulk operations unless explicitly requested.

**12. BATCH OVER REPEATED CALLS**
Use batch_modify_emails for bulk operations.

**13. MEETING TRANSCRIPT WORKFLOW**
- For requests like "summary of my meetings" or "show my transcriptions", call list_meeting_transcriptions first.
- Default transcript discovery to sharedWithMe=true unless the user asks otherwise.
- When the user chooses a file, ask whether they want:
  1) a summary (summarize_meeting_transcription), or
  2) the file link (open_meeting_transcription_file),
  unless they already specified one.
- Always include the source document link in transcript summary responses.

**14. SERVICE & TOOL DISCIPLINE**
Never cross service tools (Sheets ≠ Docs, GitHub ≠ Drive). Verify success from tool output.

---

## SERVICE QUICK REFERENCE

**Gmail**
search_emails supports full Gmail query syntax: from:, to:, subject:, is:unread, has:attachment, label:, etc.

**Calendar**
- Date ranges: list_events with timeMin/timeMax.
- Google Meet: create_meet_event or create_event with createMeetLink: true.
- Timezone: always pass timeZone. See DATE CONTEXT section for details.
- Attendee names: pass the name directly; the system resolves emails from Gmail history.
- "Just me" / "only me" → self-only event, never resolve to another contact.
- If check_person_availability fails, you can still add them as attendee. Calendar sends invites automatically.
- get_event requires a Calendar eventId — not a Gmail messageId.

**Email attachments (user-uploaded files)**
Use the attachments parameter with the localPath provided in the attached files list. Example: attachments: [{ localPath: "/path/to/file" }]

**Drive**
list_drive_files → share_drive_file / download_drive_file / extract_drive_file_text / convert_file_to_google_doc. append_drive_document_text to append.

**Sheets**
list_spreadsheets → read_sheet_values before edits. update_timesheet_hours for date-based entries (never hardcode rows). MCP fallback prefix: ${SHEETS_MCP_TOOL_PREFIX}.

**Google Docs**
list_documents → get_document_text → create_document / insert_text / append_text / replace_text. No write scope? Use append_drive_document_text.

**Meeting Transcription**
Use list_meeting_transcriptions to discover transcript docs (sharedWithMe by default), then:
- summarize_meeting_transcription for structured notes, or
- open_meeting_transcription_file for direct doc link.

**Google Chat**
list_chat_spaces → send_chat_message.

**GitHub**
Use owner/repo format. search_repos for discovery.
- **Undo/Revert**: Use \`revert_commit\` to safely undo changes by adding a new commit.
- **Permanently Remove**: Use \`reset_branch\` ONLY if the user explicitly asks to "remove", "delete", or "hard reset" a commit PERMANENTLY. This is destructive and rewrites history.

**Outlook**
Prefix: outlook_. KQL search: from:, subject:, hasAttachment:true. Use outlook_list_folders before moving.

**Teams**
Prefix: teams_. teams_list_teams → teams_list_channels → teams_send_channel_message. Chats: teams_list_chats → teams_send_chat_message.

**GCS**
Prefix: gcs_. Bucket names are lowercase-hyphen. gcs_list_buckets → gcs_list_objects (use prefix for folders). gcs_batch_delete_objects requires confirmDelete: true. gcs_generate_signed_url for temp links.

**Service disambiguation**
Lowercase-hyphen names (e.g. "my-data-bucket") → try GCS first, then Drive.

---

## RESPONSE QUALITY
- Summarise concisely. Include links/IDs the user needs. State failures clearly. No chain-of-thought.`;

    const dateContextPrompt = `DATE CONTEXT
Now: ${dateContext.nowIso} | TZ: ${dateContext.timeZone} | Today: ${dateContext.today} (${dateContext.weekday}) | Tomorrow: ${dateContext.tomorrow} | Yesterday: ${dateContext.yesterday}

- Resolve relative dates (today, tomorrow, etc.) from this context.
- Calendar lookups: always send explicit ISO timeMin/timeMax.
- Calendar events: use the user's spoken time directly in ISO (do NOT convert to UTC). Pass timeZone separately. E.g. "3pm IST" → startDateTime: "2026-02-11T15:00:00", timeZone: "Asia/Kolkata".`;

    return `${basePrompt}\n\n${dateContextPrompt}`;
}

// Streaming version with real-time callbacks
async function runAgentConversationStreaming({ message, history = [], attachedFiles = [], onTextChunk, onToolStart, onToolEnd }) {
    if (!process.env.OPENAI_API_KEY) {
        throw new Error('OpenAI API key missing');
    }

    const emailSendConfirmedForTurn = isEmailSendConfirmedForTurn({ message, history });

    const normalizeAssistantText = (content) => {
        if (typeof content === 'string') return content.trim();
        if (Array.isArray(content)) {
            return content.map(part => {
                if (typeof part === 'string') return part;
                if (part && typeof part.text === 'string') return part.text;
                return '';
            }).join('\n').trim();
        }
        return '';
    };

    const safeJsonStringify = (value) => {
        try {
            return JSON.stringify(value);
        } catch (error) {
            return JSON.stringify({ error: 'Unable to serialize tool output', message: String(error?.message || error) });
        }
    };

    const compactValueForModel = (value, depth = 0) => {
        if (value === null || value === undefined) return value;
        if (typeof value === 'string') {
            return truncateText(value, MODEL_TOOL_VALUE_MAX_STRING_CHARS);
        }
        if (typeof value === 'number' || typeof value === 'boolean') return value;
        if (Array.isArray(value)) {
            if (depth >= 4) return `[array(${value.length}) truncated]`;
            const capped = value
                .slice(0, MODEL_TOOL_VALUE_MAX_ARRAY_ITEMS)
                .map(item => compactValueForModel(item, depth + 1));
            if (value.length > MODEL_TOOL_VALUE_MAX_ARRAY_ITEMS) {
                capped.push(`[${value.length - MODEL_TOOL_VALUE_MAX_ARRAY_ITEMS} more items truncated]`);
            }
            return capped;
        }
        if (typeof value === 'object') {
            if (depth >= 4) return '[object truncated]';
            const entries = Object.entries(value);
            const limitedEntries = entries.slice(0, MODEL_TOOL_VALUE_MAX_OBJECT_KEYS);
            const out = {};
            for (const [key, val] of limitedEntries) {
                out[key] = compactValueForModel(val, depth + 1);
            }
            if (entries.length > MODEL_TOOL_VALUE_MAX_OBJECT_KEYS) {
                out.__truncatedKeys = entries.length - MODEL_TOOL_VALUE_MAX_OBJECT_KEYS;
            }
            return out;
        }
        return truncateText(String(value), MODEL_TOOL_VALUE_MAX_STRING_CHARS);
    };

    const compactToolPayloadForModel = (payload) => {
        const compacted = compactValueForModel(payload);
        const json = safeJsonStringify(compacted);
        if (json.length <= MODEL_TOOL_RESULT_MAX_CHARS) {
            return compacted;
        }
        return {
            summary: 'Tool output was too large and was compacted for model context.',
            preview: truncateText(json, MODEL_TOOL_RESULT_MAX_CHARS)
        };
    };

    const parseToolCallArguments = (rawArguments) => {
        if (rawArguments && typeof rawArguments === 'object' && !Array.isArray(rawArguments)) {
            return rawArguments;
        }
        const raw = String(rawArguments || '').trim();
        if (!raw) return {};
        const candidates = [];
        const pushCandidate = (value) => {
            const normalized = String(value || '').trim();
            if (normalized && !candidates.includes(normalized)) candidates.push(normalized);
        };
        pushCandidate(raw);
        const fenced = raw.match(/^```(?:json)?\s*([\s\S]*?)\s*```$/i);
        if (fenced && fenced[1]) pushCandidate(fenced[1]);
        const latest = candidates[candidates.length - 1] || raw;
        const firstBrace = latest.search(/[{\[]/);
        const lastBrace = Math.max(latest.lastIndexOf('}'), latest.lastIndexOf(']'));
        if (firstBrace >= 0 && lastBrace > firstBrace) {
            pushCandidate(latest.slice(firstBrace, lastBrace + 1));
        }
        const quoteNormalized = (candidates[candidates.length - 1] || raw)
            .replace(/[""]/g, '"')
            .replace(/['']/g, "'");
        pushCandidate(quoteNormalized);
        pushCandidate(quoteNormalized.replace(/,\s*([}\]])/g, '$1'));
        let lastError = null;
        for (const candidate of candidates) {
            try {
                const parsed = JSON.parse(candidate);
                if (!parsed || typeof parsed !== 'object' || Array.isArray(parsed)) {
                    throw new Error('Tool arguments must be a JSON object.');
                }
                return parsed;
            } catch (error) {
                lastError = error;
            }
        }
        throw new Error(`Invalid arguments JSON: ${lastError ? lastError.message : 'Unable to parse tool arguments'}`);
    };

    const toolContext = getConnectedToolContext();
    const { selectedTools, routingHints } = selectToolsForMessageIntent({
        userMessage: message,
        availableTools: toolContext.availableTools,
        docsConnected: toolContext.docsConnected,
        docsListOnlyConnected: toolContext.docsListOnlyConnected
    });
    const requestTools = cloneToolsForRequest(selectedTools);
    const dateContext = getCurrentDateContext();

    // Get user's first name for email signatures (injected into email tool descriptions only)
    let userFirstName = null;
    try {
        const userEmail = await getPrimaryEmailAddress();
        if (userEmail) {
            const localPart = userEmail.split('@')[0];
            const namePart = localPart.split(/[._\-+0-9]/)[0];
            if (namePart && namePart.length >= 2) {
                userFirstName = namePart.charAt(0).toUpperCase() + namePart.slice(1).toLowerCase();
            }
        }
    } catch (err) {
        console.log('Could not get user name for signature:', err.message);
    }
    applyEmailSignatureHintToTools(requestTools, userFirstName);

    const systemPrompt = buildAgentSystemPrompt({
        statusText: toolContext.statusText,
        toolCount: requestTools.length,
        dateContext,
        connectedServices: toolContext.connectedServices
    });
    const routingHintBlock = routingHints.length > 0
        ? `\n\n[Tool Routing Hints]\n${routingHints.map(hint => `- ${hint}`).join('\n')}`
        : '';

    const attachedFilesBlock = attachedFiles.length > 0
        ? `\n\n[Attached Files]\nThe user has attached ${attachedFiles.length} file(s) to this message:\n${attachedFiles.map((f, i) => `${i + 1}. "${f.name}" (${(f.size / 1024).toFixed(1)} KB, ${f.mimeType})\n   - File ID: ${f.fileId}\n   - Local path: ${f.localPath}`).join('\n')}\n\n**IMPORTANT**: To upload these files, use the localPath parameter:\n- Upload to Drive: create_drive_file({ name: "filename", localPath: "path/from/above" })\n- Upload to GCS: gcs_upload_object({ bucket: "bucket-name", name: "filename", localPath: "path/from/above" })\n- For emails: Read file content from localPath and encode as base64 for attachment\n- DO NOT pass localPath as content - use it as the localPath parameter!`
        : '';

    const enrichedUserMessage = `${message}
${routingHintBlock}
${attachedFilesBlock}

[Runtime Date Context]
- Current timestamp (UTC): ${dateContext.nowIso}
- Local timezone: ${dateContext.timeZone}
- Today: ${dateContext.today}
- Tomorrow: ${dateContext.tomorrow}
- Yesterday: ${dateContext.yesterday}`;

    const safeHistory = buildSafeChatHistory(history);
    const messages = [
        { role: 'system', content: systemPrompt },
        ...safeHistory,
        { role: 'user', content: enrichedUserMessage }
    ];
    const toolChoice = requestTools.length > 0 ? 'auto' : undefined;

    // Use OpenAI streaming for the first call
    const request = {
        model: OPENAI_MODEL,
        messages,
        temperature: OPENAI_TEMPERATURE,
        stream: true
    };
    if (Array.isArray(requestTools) && requestTools.length > 0) {
        request.tools = requestTools;
        request.tool_choice = toolChoice || 'auto';
        request.parallel_tool_calls = true;
    }
    if (OPENAI_MAX_OUTPUT_TOKENS) {
        request.max_tokens = OPENAI_MAX_OUTPUT_TOKENS;
    }

    const stream = await openai.chat.completions.create(request);
    let assistantMessage = { role: 'assistant', content: '', tool_calls: [] };
    let currentToolCall = null;

    // Process stream
    for await (const chunk of stream) {
        const delta = chunk.choices?.[0]?.delta;
        if (!delta) continue;

        // Stream text content
        if (delta.content) {
            assistantMessage.content += delta.content;
            if (onTextChunk) onTextChunk(delta.content);
        }

        // Collect tool calls
        if (delta.tool_calls) {
            for (const toolDelta of delta.tool_calls) {
                if (toolDelta.index !== undefined) {
                    if (!assistantMessage.tool_calls[toolDelta.index]) {
                        assistantMessage.tool_calls[toolDelta.index] = {
                            id: toolDelta.id || '',
                            type: 'function',
                            function: { name: '', arguments: '' }
                        };
                    }
                    const toolCall = assistantMessage.tool_calls[toolDelta.index];
                    if (toolDelta.id) toolCall.id = toolDelta.id;
                    if (toolDelta.function?.name) toolCall.function.name += toolDelta.function.name;
                    if (toolDelta.function?.arguments) toolCall.function.arguments += toolDelta.function.arguments;
                }
            }
        }
    }

    // Clean up tool_calls array (remove empty entries)
    assistantMessage.tool_calls = assistantMessage.tool_calls.filter(tc => tc && tc.function?.name);

    const allToolResults = [];
    const allSteps = [];
    const MAX_TURNS = 15;
    let turnCount = 0;
    let activeModel = OPENAI_MODEL;

    // Continue with tool execution loop (non-streaming for tool calls)
    while (assistantMessage.tool_calls && assistantMessage.tool_calls.length > 0 && turnCount < MAX_TURNS) {
        turnCount += 1;
        messages.push(assistantMessage);

        const toolPromises = assistantMessage.tool_calls.map(async (toolCall) => {
            const toolName = toolCall.function.name;
            let args;
            try {
                args = parseToolCallArguments(toolCall.function.arguments);
            } catch (error) {
                return { toolCall, error: `Invalid arguments: ${error.message}` };
            }

            if (EMAIL_SEND_CONFIRMATION_TOOLS.has(toolName)) {
                try {
                    args = await prepareEmailSendArgs(toolName, args || {});
                } catch (error) {
                    const prepError = error.message;
                    if (onToolEnd) onToolEnd(toolName, prepError, false, turnCount);
                    return { toolCall, error: prepError };
                }
            }
            if (EMAIL_SEND_CONFIRMATION_TOOLS.has(toolName) && !emailSendConfirmedForTurn) {
                const confirmationError = buildEmailSendConfirmationMessage(toolName, args || {});
                if (onToolEnd) onToolEnd(toolName, confirmationError, false, turnCount);
                return { toolCall, error: confirmationError };
            }
            if (EMAIL_SEND_CONFIRMATION_TOOLS.has(toolName)) {
                args = { ...(args || {}), confirmSend: true };
            }

            args = applyUserIntentGuardsToToolArgs({
                toolName,
                args,
                userMessage: message
            });

            if (onToolStart) onToolStart(toolName, args, turnCount);
            const step = { tool: toolName, args, turn: turnCount, timestamp: Date.now() };
            try {
                const result = await executeTool(toolName, args);
                step.result = result;
                step.success = true;
                if (onToolEnd) onToolEnd(toolName, result, true, turnCount);
                return { toolCall, result, step };
            } catch (error) {
                step.error = error.message;
                step.success = false;
                if (onToolEnd) onToolEnd(toolName, error.message, false, turnCount);
                return { toolCall, error: error.message, step };
            }
        });

        const toolResults = await Promise.all(toolPromises);
        for (const { toolCall, result, error, step } of toolResults) {
            if (step) allSteps.push(step);
            allToolResults.push(error
                ? { tool: toolCall.function.name, error }
                : { tool: toolCall.function.name, result: compactValueForModel(result) });

            const toolPayloadForModel = compactToolPayloadForModel(error ? { error } : result);
            messages.push({
                role: 'tool',
                tool_call_id: toolCall.id,
                content: safeJsonStringify(toolPayloadForModel)
            });
        }

        const userInputRequiredResult = toolResults.find(item => item?.error && shouldRequestUserInputForToolError(item.error));
        if (userInputRequiredResult) {
            return {
                response: userInputRequiredResult.error,
                toolResults: allToolResults,
                steps: allSteps,
                turnsUsed: turnCount,
                model: activeModel
            };
        }

        // Next turn - use streaming again
        const nextRequest = {
            model: OPENAI_MODEL,
            messages,
            temperature: OPENAI_TEMPERATURE,
            stream: true
        };
        if (Array.isArray(requestTools) && requestTools.length > 0) {
            nextRequest.tools = requestTools;
            nextRequest.tool_choice = toolChoice || 'auto';
            nextRequest.parallel_tool_calls = true;
        }
        if (OPENAI_MAX_OUTPUT_TOKENS) {
            nextRequest.max_tokens = OPENAI_MAX_OUTPUT_TOKENS;
        }

        const nextStream = await openai.chat.completions.create(nextRequest);
        assistantMessage = { role: 'assistant', content: '', tool_calls: [] };

        for await (const chunk of nextStream) {
            const delta = chunk.choices?.[0]?.delta;
            if (!delta) continue;
            if (delta.content) {
                assistantMessage.content += delta.content;
                if (onTextChunk) onTextChunk(delta.content);
            }
            if (delta.tool_calls) {
                for (const toolDelta of delta.tool_calls) {
                    if (toolDelta.index !== undefined) {
                        if (!assistantMessage.tool_calls[toolDelta.index]) {
                            assistantMessage.tool_calls[toolDelta.index] = {
                                id: toolDelta.id || '',
                                type: 'function',
                                function: { name: '', arguments: '' }
                            };
                        }
                        const toolCall = assistantMessage.tool_calls[toolDelta.index];
                        if (toolDelta.id) toolCall.id = toolDelta.id;
                        if (toolDelta.function?.name) toolCall.function.name += toolDelta.function.name;
                        if (toolDelta.function?.arguments) toolCall.function.arguments += toolDelta.function.arguments;
                    }
                }
            }
        }

        assistantMessage.tool_calls = assistantMessage.tool_calls.filter(tc => tc && tc.function?.name);
    }

    let finalResponse = normalizeAssistantText(assistantMessage.content);
    if (turnCount >= MAX_TURNS && assistantMessage.tool_calls && assistantMessage.tool_calls.length > 0) {
        finalResponse += '\n\n(Reached maximum steps for this request. Some operations may still be pending.)';
    }
    if (!finalResponse && allToolResults.length > 0) {
        finalResponse = 'Completed the requested workflow. See the tool results for details.';
    }
    finalResponse = truncateText(finalResponse, ASSISTANT_RESPONSE_MAX_CHARS);

    return {
        response: finalResponse,
        toolResults: allToolResults,
        steps: allSteps,
        turnsUsed: turnCount,
        model: activeModel
    };
}

async function runAgentConversation({ message, history = [], attachedFiles = [] }) {
    if (!process.env.OPENAI_API_KEY) {
        throw new Error('OpenAI API key missing');
    }

    const emailSendConfirmedForTurn = isEmailSendConfirmedForTurn({ message, history });

    const normalizeAssistantText = (content) => {
        if (typeof content === 'string') return content.trim();
        if (Array.isArray(content)) {
            return content
                .map(part => {
                    if (typeof part === 'string') return part;
                    if (part && typeof part.text === 'string') return part.text;
                    return '';
                })
                .join('\n')
                .trim();
        }
        return '';
    };

    const safeJsonStringify = (value) => {
        try {
            return JSON.stringify(value);
        } catch (error) {
            return JSON.stringify({
                error: 'Unable to serialize tool output',
                message: String(error?.message || error)
            });
        }
    };

    const compactValueForModel = (value, depth = 0) => {
        if (value === null || value === undefined) return value;
        if (typeof value === 'string') {
            return truncateText(value, MODEL_TOOL_VALUE_MAX_STRING_CHARS);
        }
        if (typeof value === 'number' || typeof value === 'boolean') return value;
        if (Array.isArray(value)) {
            if (depth >= 4) return `[array(${value.length}) truncated]`;
            const capped = value
                .slice(0, MODEL_TOOL_VALUE_MAX_ARRAY_ITEMS)
                .map(item => compactValueForModel(item, depth + 1));
            if (value.length > MODEL_TOOL_VALUE_MAX_ARRAY_ITEMS) {
                capped.push(`[${value.length - MODEL_TOOL_VALUE_MAX_ARRAY_ITEMS} more items truncated]`);
            }
            return capped;
        }
        if (typeof value === 'object') {
            if (depth >= 4) return '[object truncated]';
            const entries = Object.entries(value);
            const limitedEntries = entries.slice(0, MODEL_TOOL_VALUE_MAX_OBJECT_KEYS);
            const out = {};
            for (const [key, val] of limitedEntries) {
                out[key] = compactValueForModel(val, depth + 1);
            }
            if (entries.length > MODEL_TOOL_VALUE_MAX_OBJECT_KEYS) {
                out.__truncatedKeys = entries.length - MODEL_TOOL_VALUE_MAX_OBJECT_KEYS;
            }
            return out;
        }
        return truncateText(String(value), MODEL_TOOL_VALUE_MAX_STRING_CHARS);
    };

    const compactToolPayloadForModel = (payload) => {
        const compacted = compactValueForModel(payload);
        const json = safeJsonStringify(compacted);
        if (json.length <= MODEL_TOOL_RESULT_MAX_CHARS) {
            return compacted;
        }
        return {
            summary: 'Tool output was too large and was compacted for model context.',
            preview: truncateText(json, MODEL_TOOL_RESULT_MAX_CHARS)
        };
    };

    const parseToolCallArguments = (rawArguments) => {
        if (rawArguments && typeof rawArguments === 'object' && !Array.isArray(rawArguments)) {
            return rawArguments;
        }
        const raw = String(rawArguments || '').trim();
        if (!raw) return {};

        const candidates = [];
        const pushCandidate = (value) => {
            const normalized = String(value || '').trim();
            if (normalized && !candidates.includes(normalized)) candidates.push(normalized);
        };

        pushCandidate(raw);
        const fenced = raw.match(/^```(?:json)?\s*([\s\S]*?)\s*```$/i);
        if (fenced && fenced[1]) pushCandidate(fenced[1]);

        const latest = candidates[candidates.length - 1] || raw;
        const firstBrace = latest.search(/[{\[]/);
        const lastBrace = Math.max(latest.lastIndexOf('}'), latest.lastIndexOf(']'));
        if (firstBrace >= 0 && lastBrace > firstBrace) {
            pushCandidate(latest.slice(firstBrace, lastBrace + 1));
        }

        const quoteNormalized = (candidates[candidates.length - 1] || raw)
            .replace(/[“”]/g, '"')
            .replace(/[‘’]/g, "'");
        pushCandidate(quoteNormalized);
        pushCandidate(quoteNormalized.replace(/,\s*([}\]])/g, '$1'));

        let lastError = null;
        for (const candidate of candidates) {
            try {
                const parsed = JSON.parse(candidate);
                if (!parsed || typeof parsed !== 'object' || Array.isArray(parsed)) {
                    throw new Error('Tool arguments must be a JSON object.');
                }
                return parsed;
            } catch (error) {
                lastError = error;
            }
        }

        throw new Error(`Invalid arguments JSON: ${lastError ? lastError.message : 'Unable to parse tool arguments'}`);
    };

    const isRetriableOpenAiError = (error) => {
        const status = Number(error?.status || error?.code || error?.response?.status || 0);
        if ([408, 409, 425, 429, 500, 502, 503, 504].includes(status)) return true;
        const code = String(error?.code || '').toUpperCase();
        if (['ETIMEDOUT', 'ECONNRESET', 'EAI_AGAIN', 'ENOTFOUND'].includes(code)) return true;
        const message = String(error?.message || '');
        return /rate limit|timeout|timed out|temporar|overloaded|network|try again/i.test(message);
    };

    const buildOpenAiRequest = ({ model, messages, tools, toolChoice }) => {
        const request = {
            model,
            messages,
            temperature: OPENAI_TEMPERATURE
        };
        if (Array.isArray(tools) && tools.length > 0) {
            request.tools = tools;
            request.tool_choice = toolChoice || 'auto';
            request.parallel_tool_calls = true;
        }
        if (OPENAI_MAX_OUTPUT_TOKENS) {
            request.max_tokens = OPENAI_MAX_OUTPUT_TOKENS;
        }
        return request;
    };

    const createCompletionWithRetry = async ({ messages, tools, toolChoice }) => {
        const modelCandidates = [OPENAI_MODEL, OPENAI_FALLBACK_MODEL]
            .map(value => String(value || '').trim())
            .filter(Boolean)
            .filter((value, index, list) => list.indexOf(value) === index);
        let lastError = null;

        for (const model of modelCandidates) {
            for (let attempt = 0; attempt <= OPENAI_CHAT_MAX_RETRIES; attempt += 1) {
                try {
                    const response = await openai.chat.completions.create(
                        buildOpenAiRequest({ model, messages, tools, toolChoice })
                    );
                    return { response, model };
                } catch (error) {
                    lastError = error;
                    const canRetry = isRetriableOpenAiError(error) && attempt < OPENAI_CHAT_MAX_RETRIES;
                    if (!canRetry) break;
                    const delayMs = Math.min(3000, 350 * (2 ** attempt));
                    await sleep(delayMs);
                }
            }
        }

        throw lastError || new Error('OpenAI completion failed');
    };

    const toolContext = getConnectedToolContext();
    const { selectedTools, routingHints } = selectToolsForMessageIntent({
        userMessage: message,
        availableTools: toolContext.availableTools,
        docsConnected: toolContext.docsConnected,
        docsListOnlyConnected: toolContext.docsListOnlyConnected
    });
    const requestTools = cloneToolsForRequest(selectedTools);
    const dateContext = getCurrentDateContext();

    // Get user's first name for email signatures (injected into email tool descriptions only)
    let userFirstName = null;
    try {
        const userEmail = await getPrimaryEmailAddress();
        if (userEmail) {
            const localPart = userEmail.split('@')[0];
            const namePart = localPart.split(/[._\-+0-9]/)[0];
            if (namePart && namePart.length >= 2) {
                userFirstName = namePart.charAt(0).toUpperCase() + namePart.slice(1).toLowerCase();
            }
        }
    } catch (err) {
        console.log('Could not get user name for signature:', err.message);
    }
    applyEmailSignatureHintToTools(requestTools, userFirstName);

    const systemPrompt = buildAgentSystemPrompt({
        statusText: toolContext.statusText,
        toolCount: requestTools.length,
        dateContext,
        connectedServices: toolContext.connectedServices
    });
    const routingHintBlock = routingHints.length > 0
        ? `\n\n[Tool Routing Hints]\n${routingHints.map(hint => `- ${hint}`).join('\n')}`
        : '';

    const attachedFilesBlock = attachedFiles.length > 0
        ? `\n\n[Attached Files]\nThe user has attached ${attachedFiles.length} file(s) to this message:\n${attachedFiles.map((f, i) => `${i + 1}. "${f.name}" (${(f.size / 1024).toFixed(1)} KB, ${f.mimeType})\n   - File ID: ${f.fileId}\n   - Local path: ${f.localPath}`).join('\n')}\n\n**IMPORTANT**: To upload these files, use the localPath parameter:\n- Upload to Drive: create_drive_file({ name: "filename", localPath: "path/from/above" })\n- Upload to GCS: gcs_upload_object({ bucket: "bucket-name", name: "filename", localPath: "path/from/above" })\n- For emails: Read file content from localPath and encode as base64 for attachment\n- DO NOT pass localPath as content - use it as the localPath parameter!`
        : '';

    const enrichedUserMessage = `${message}
${routingHintBlock}
${attachedFilesBlock}

[Runtime Date Context]
- Current timestamp (UTC): ${dateContext.nowIso}
- Local timezone: ${dateContext.timeZone}
- Today: ${dateContext.today}
- Tomorrow: ${dateContext.tomorrow}
- Yesterday: ${dateContext.yesterday}`;

    const safeHistory = buildSafeChatHistory(history);
    const messages = [
        { role: 'system', content: systemPrompt },
        ...safeHistory,
        { role: 'user', content: enrichedUserMessage }
    ];
    const toolChoice = requestTools.length > 0 ? 'auto' : undefined;

    let completion = await createCompletionWithRetry({
        messages,
        tools: requestTools,
        toolChoice
    });
    let activeModel = completion.model;
    let assistantMessage = completion.response?.choices?.[0]?.message;
    if (!assistantMessage) {
        throw new Error('OpenAI returned no assistant message.');
    }

    const allToolResults = [];
    const allSteps = [];
    const MAX_TURNS = 15;
    let turnCount = 0;

    while (assistantMessage.tool_calls && turnCount < MAX_TURNS) {
        turnCount += 1;
        messages.push(assistantMessage);

        const toolPromises = assistantMessage.tool_calls.map(async (toolCall) => {
            const toolName = toolCall.function.name;
            let args;
            try {
                args = parseToolCallArguments(toolCall.function.arguments);
            } catch (error) {
                return { toolCall, error: `Invalid arguments: ${error.message}` };
            }

            if (EMAIL_SEND_CONFIRMATION_TOOLS.has(toolName)) {
                try {
                    args = await prepareEmailSendArgs(toolName, args || {});
                } catch (error) {
                    return { toolCall, error: error.message };
                }
            }
            if (EMAIL_SEND_CONFIRMATION_TOOLS.has(toolName) && !emailSendConfirmedForTurn) {
                return { toolCall, error: buildEmailSendConfirmationMessage(toolName, args || {}) };
            }
            if (EMAIL_SEND_CONFIRMATION_TOOLS.has(toolName)) {
                args = { ...(args || {}), confirmSend: true };
            }

            args = applyUserIntentGuardsToToolArgs({
                toolName,
                args,
                userMessage: message
            });

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
            allToolResults.push(error
                ? { tool: toolCall.function.name, error }
                : { tool: toolCall.function.name, result: compactValueForModel(result) });

            const toolPayloadForModel = compactToolPayloadForModel(error ? { error } : result);
            messages.push({
                role: 'tool',
                tool_call_id: toolCall.id,
                content: safeJsonStringify(toolPayloadForModel)
            });
        }

        const userInputRequiredResult = toolResults.find(item => item?.error && shouldRequestUserInputForToolError(item.error));
        if (userInputRequiredResult) {
            return {
                response: userInputRequiredResult.error,
                toolResults: allToolResults,
                steps: allSteps,
                turnsUsed: turnCount,
                model: activeModel
            };
        }

        completion = await createCompletionWithRetry({
            messages,
            tools: requestTools,
            toolChoice
        });
        activeModel = completion.model;
        assistantMessage = completion.response?.choices?.[0]?.message;
        if (!assistantMessage) {
            throw new Error('OpenAI returned no assistant follow-up message.');
        }
    }

    let finalResponse = normalizeAssistantText(assistantMessage.content);
    if (turnCount >= MAX_TURNS && assistantMessage.tool_calls) {
        finalResponse += '\n\n(Reached maximum steps for this request. Some operations may still be pending.)';
    }
    if (!finalResponse && allToolResults.length > 0) {
        finalResponse = 'Completed the requested workflow. See the tool results for details.';
    }
    finalResponse = truncateText(finalResponse, ASSISTANT_RESPONSE_MAX_CHARS);

    return {
        response: finalResponse,
        toolResults: allToolResults,
        steps: allSteps,
        turnsUsed: turnCount,
        model: activeModel
    };
}

const runningScheduledTaskIds = new Set();

async function runScheduledTask(taskId, trigger = 'scheduled') {
    const task = scheduledTasks.find(item => item.id === taskId);
    if (!task) throw new Error('Task not found');
    if (runningScheduledTaskIds.has(taskId)) {
        return { success: false, skipped: true, message: 'Task is already running.' };
    }

    runningScheduledTaskIds.add(taskId);
    const now = schedulerDateParts(new Date());
    try {
        const scheduledMessage = `${task.instruction}

[Scheduled Task Context]
- Trigger: ${trigger}
- Task Name: ${task.name}
- Scheduled Time: ${task.time}
- Execution Timestamp (UTC): ${now.timestamp}
- Run this task fully now.`;

        const result = await runAgentConversation({ message: scheduledMessage, history: [] });
        task.lastRunAt = now.timestamp;
        task.lastRunDate = now.date;
        task.lastStatus = 'success';
        task.lastError = '';
        task.lastResponse = String(result.response || '').slice(0, 1500);
        task.updatedAt = new Date().toISOString();
        saveScheduledTasksToDisk();
        return { success: true, taskId, result };
    } catch (error) {
        task.lastRunAt = now.timestamp;
        task.lastRunDate = now.date;
        task.lastStatus = 'failed';
        task.lastError = error.message;
        task.updatedAt = new Date().toISOString();
        saveScheduledTasksToDisk();
        return { success: false, taskId, error: error.message };
    } finally {
        runningScheduledTaskIds.delete(taskId);
    }
}

async function runSchedulerTick() {
    if (schedulerTickInProgress) return;
    schedulerTickInProgress = true;
    try {
        const now = schedulerDateParts(new Date());
        const nowMinutes = dailyTimeToMinutes(now.hhmm);
        const dueTasks = scheduledTasks
            .filter(task => {
                if (!task.enabled || task.lastRunDate === now.date) return false;
                const scheduledMinutes = dailyTimeToMinutes(task.time);
                if (scheduledMinutes === null || nowMinutes === null) return false;
                return scheduledMinutes <= nowMinutes;
            })
            .sort((a, b) => a.time.localeCompare(b.time));
        for (const task of dueTasks) {
            const result = await runScheduledTask(task.id, 'scheduled');
            if (!result.success && !result.skipped) {
                console.error(`Scheduled task "${task.name}" failed: ${result.error || result.message}`);
            } else if (result.success) {
                console.log(`Scheduled task "${task.name}" executed at ${now.hhmm}`);
            }
        }
    } finally {
        schedulerTickInProgress = false;
    }
}

function startScheduledTaskRunner() {
    if (schedulerInterval) clearInterval(schedulerInterval);
    schedulerInterval = setInterval(() => {
        runSchedulerTick().catch(error => {
            console.error('Scheduler tick failed:', error.message);
        });
    }, 30000);
    runSchedulerTick().catch(error => {
        console.error('Initial scheduler tick failed:', error.message);
    });
}

// Get single email content
app.get('/api/gmail/message/:id', async (req, res) => {
    try {
        if (!gmailClient) {
            return res.status(401).json({ error: 'Gmail not connected' });
        }
        const { id } = req.params;
        const email = await readEmail({ messageId: id });
        res.json(email);
    } catch (error) {
        console.error('Error fetching email:', error);
        res.status(500).json({ error: 'Internal server error' });
    }
});

// File upload endpoint
app.post('/api/upload', upload.array('files', 10), (req, res) => {
    try {
        if (!req.files || req.files.length === 0) {
            return res.status(400).json({ error: 'No files uploaded' });
        }

        const uploadedFileInfos = req.files.map(file => {
            const fileId = crypto.randomBytes(16).toString('hex');
            const fileInfo = {
                fileId,
                path: file.path,
                originalName: file.originalname,
                size: file.size,
                mimeType: file.mimetype,
                uploadedAt: Date.now()
            };
            uploadedFiles.set(fileId, fileInfo);

            return {
                fileId,
                name: file.originalname,
                size: file.size,
                mimeType: file.mimetype
            };
        });

        res.json({ files: uploadedFileInfos });
    } catch (error) {
        console.error('Upload error:', error);
        res.status(500).json({ error: 'Failed to upload files' });
    }
});

// Get uploaded file info
app.get('/api/upload/:fileId', (req, res) => {
    const fileId = req.params.fileId;
    const fileInfo = uploadedFiles.get(fileId);

    if (!fileInfo) {
        return res.status(404).json({ error: 'File not found' });
    }

    res.json({
        fileId,
        name: fileInfo.originalName,
        size: fileInfo.size,
        mimeType: fileInfo.mimeType
    });
});

// Delete uploaded file
app.delete('/api/upload/:fileId', (req, res) => {
    const fileId = req.params.fileId;
    const fileInfo = uploadedFiles.get(fileId);

    if (!fileInfo) {
        return res.status(404).json({ error: 'File not found' });
    }

    try {
        if (fs.existsSync(fileInfo.path)) {
            fs.unlinkSync(fileInfo.path);
        }
        uploadedFiles.delete(fileId);
        res.json({ success: true });
    } catch (error) {
        console.error('Failed to delete file:', error);
        res.status(500).json({ error: 'Failed to delete file' });
    }
});

// Streaming chat endpoint (Server-Sent Events)
app.post('/api/chat/stream', async (req, res) => {
    const { message, history = [], attachedFiles = [] } = req.body;
    try {
        if (typeof message !== 'string' || !message.trim()) {
            return res.status(400).json({ error: 'message must be a non-empty string' });
        }
        if (message.length > MAX_CHAT_MESSAGE_CHARS) {
            return res.status(400).json({ error: `message exceeds max length (${MAX_CHAT_MESSAGE_CHARS} chars)` });
        }
        if (Array.isArray(history) && history.length > MAX_CHAT_HISTORY_ITEMS) {
            return res.status(400).json({ error: `history exceeds max items (${MAX_CHAT_HISTORY_ITEMS})` });
        }
        if (Array.isArray(attachedFiles) && attachedFiles.length > MAX_CHAT_ATTACHED_FILES) {
            return res.status(400).json({ error: `too many attached files (max ${MAX_CHAT_ATTACHED_FILES})` });
        }

        // Set up SSE headers
        res.setHeader('Content-Type', 'text/event-stream');
        res.setHeader('Cache-Control', 'no-cache');
        res.setHeader('Connection', 'keep-alive');
        res.flushHeaders();

        const sendEvent = (event, data) => {
            res.write(`event: ${event}\ndata: ${JSON.stringify(data)}\n\n`);
        };

        const safeHistory = buildSafeChatHistory(history);

        // Process attached files
        const fileContexts = [];
        if (Array.isArray(attachedFiles) && attachedFiles.length > 0) {
            for (const fileId of attachedFiles) {
                const fileInfo = uploadedFiles.get(fileId);
                if (fileInfo) {
                    fileContexts.push({
                        fileId,
                        name: fileInfo.originalName,
                        size: fileInfo.size,
                        mimeType: fileInfo.mimeType,
                        localPath: fileInfo.path
                    });
                }
            }
        }

        // Run agent with streaming callbacks
        const result = await runAgentConversationStreaming({
            message,
            history: safeHistory,
            attachedFiles: fileContexts,
            onTextChunk: (chunk) => sendEvent('text', { chunk }),
            onToolStart: (tool, args, turn) => sendEvent('tool_start', { tool, args, turn }),
            onToolEnd: (tool, result, success, turn) => sendEvent('tool_end', { tool, result, success, turn })
        });

        // Send final result
        sendEvent('done', result);
        res.end();
    } catch (error) {
        console.error('Stream chat error:', error);
        res.write(`event: error\ndata: ${JSON.stringify({ error: 'Something went wrong. Please try again.' })}\n\n`);
        res.end();
    }
});

// Google Meet endpoints

// Get meeting metadata from Calendar by meeting code
app.get('/api/meet/metadata', async (req, res) => {
    try {
        const { code } = req.query;

        if (!code) {
            return res.status(400).json({ error: 'Meeting code is required' });
        }

        if (!calendarClient || !hasCalendarScope()) {
            return res.status(503).json({
                error: 'Calendar service not available',
                setupRequired: true
            });
        }

        // Get today's events
        const today = new Date();
        const timeMin = new Date(today.setHours(0, 0, 0, 0)).toISOString();
        const timeMax = new Date(today.setHours(23, 59, 59, 999)).toISOString();

        const eventsResult = await listEvents({
            calendarId: 'primary',
            maxResults: 50,
            timeMin,
            timeMax
        });

        if (!eventsResult || !Array.isArray(eventsResult.events)) {
            return res.json({ error: 'No events found', meeting: null });
        }

        // Find event with matching Meet link
        const meeting = eventsResult.events.find(event => {
            const link = event.meetLink || event.hangoutLink || '';
            if (link) {
                return link.includes(code);
            }
            return false;
        });

        if (!meeting) {
            return res.json({
                error: 'Meeting not found in calendar',
                meeting: null
            });
        }

        // Extract attendee emails
        const attendees = (meeting.attendees || [])
            .filter(a => a.email)
            .map(a => a.email);

        res.json({
            meeting: {
                title: meeting.summary || 'Google Meet',
                meetingCode: code,
                startTime: meeting.start?.dateTime || meeting.start?.date,
                endTime: meeting.end?.dateTime || meeting.end?.date,
                attendees,
                eventId: meeting.id,
                hangoutLink: meeting.meetLink || meeting.hangoutLink || null
            }
        });

    } catch (error) {
        console.error('Error fetching meeting metadata:', error);
        res.status(500).json({ error: 'Internal server error' });
    }
});

// Finalize meeting (generate summary and create Google Doc)
app.post('/api/meet/finalize', async (req, res) => {
    try {
        const { sessionId } = req.body;

        if (!sessionId) {
            return res.status(400).json({ error: 'Session ID is required' });
        }

        // Get session from meetSessions Map
        const session = meetSessions.get(sessionId);

        if (!session) {
            return res.status(404).json({ error: 'Session not found' });
        }

        if (!docsClient || !hasDocsScope()) {
            return res.status(503).json({
                error: 'Google Docs service not available',
                setupRequired: true
            });
        }

        // Normalize and deduplicate caption stream so transcript reads cleanly.
        const transcriptNoisePatterns = [
            /arrow_downward\s*jump to bottom/gi,
            /jump to bottom/gi,
            /(^|\s)you:\s*/gi
        ];

        function sanitizeSpeakerName(rawSpeaker, fallback = 'Unknown Speaker') {
            const base = typeof rawSpeaker === 'string' ? rawSpeaker : '';
            const cleaned = base
                .replace(/\s+/g, ' ')
                .replace(/^[^A-Za-z0-9]+/, '')
                .trim();

            if (!cleaned || cleaned.length > 120) {
                return fallback;
            }

            return cleaned;
        }

        function sanitizeTranscriptText(rawText) {
            if (typeof rawText !== 'string') {
                return '';
            }

            let text = rawText.replace(/\s+/g, ' ').trim();
            for (const pattern of transcriptNoisePatterns) {
                text = text.replace(pattern, ' ');
            }

            text = text.replace(/\s+/g, ' ').trim();
            return text;
        }

        function normalizeCaptions(captions) {
            if (!Array.isArray(captions)) return [];

            return captions
                .map((caption) => {
                    if (!caption || typeof caption !== 'object') {
                        return null;
                    }

                    const speaker = sanitizeSpeakerName(caption.speaker, 'Unknown Speaker');
                    const text = sanitizeTranscriptText(caption.text);

                    if (!text) {
                        return null;
                    }

                    const parsedTimestamp = caption.timestamp ? new Date(caption.timestamp) : null;
                    const timestamp = parsedTimestamp && !Number.isNaN(parsedTimestamp.getTime())
                        ? parsedTimestamp
                        : null;

                    return { speaker, text, timestamp };
                })
                .filter(Boolean);
        }

        function deduplicateCaptions(captions) {
            const deduped = [];

            for (const caption of captions) {
                const previous = deduped[deduped.length - 1];
                if (!previous) {
                    deduped.push(caption);
                    continue;
                }

                const sameSpeaker = previous.speaker.toLowerCase() === caption.speaker.toLowerCase();
                const previousText = previous.text.toLowerCase();
                const currentText = caption.text.toLowerCase();

                // Exact repeat of prior caption update.
                if (sameSpeaker && previousText === currentText) {
                    continue;
                }

                // Streaming captions often send growing text updates; keep only final form.
                if (sameSpeaker && currentText.startsWith(previousText) && currentText.length > previousText.length) {
                    deduped[deduped.length - 1] = caption;
                    continue;
                }

                // Drop shorter fragment if full text is already kept.
                if (sameSpeaker && previousText.startsWith(currentText) && previousText.length > currentText.length) {
                    continue;
                }

                deduped.push(caption);
            }

            return deduped;
        }

        function formatCaptionTime(timestamp) {
            if (!(timestamp instanceof Date) || Number.isNaN(timestamp.getTime())) {
                return '';
            }

            return timestamp.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
        }

        function formatStepByStepTranscript(captions) {
            if (!captions || captions.length === 0) {
                return '';
            }

            return captions
                .map((caption, index) => {
                    const time = formatCaptionTime(caption.timestamp);
                    const speakerLine = `${index + 1}. ${time ? `[${time}] ` : ''}Speaker: ${caption.speaker}`;
                    return `${speakerLine}\n${caption.text}`;
                })
                .join('\n\n');
        }

        function formatSpeakerWiseTranscript(captions) {
            if (!captions || captions.length === 0) {
                return '';
            }

            const bySpeaker = new Map();

            for (const caption of captions) {
                if (!bySpeaker.has(caption.speaker)) {
                    bySpeaker.set(caption.speaker, []);
                }
                bySpeaker.get(caption.speaker).push(caption);
            }

            const sections = [];
            for (const [speaker, turns] of bySpeaker.entries()) {
                sections.push(`${speaker} (${turns.length} turn${turns.length === 1 ? '' : 's'})`);

                for (const turn of turns) {
                    const time = formatCaptionTime(turn.timestamp);
                    const prefix = time ? `[${time}] ` : '';
                    sections.push(`${prefix}${turn.text}`);
                }

                sections.push('');
            }

            return sections.join('\n').trim();
        }

        function parseJsonFromModelOutput(text) {
            const raw = String(text || '').trim();
            if (!raw) return null;

            const candidates = [raw];
            const fenced = raw.match(/```(?:json)?\s*([\s\S]*?)\s*```/i);
            if (fenced && fenced[1]) {
                candidates.push(String(fenced[1]).trim());
            }

            const firstBrace = raw.indexOf('{');
            const lastBrace = raw.lastIndexOf('}');
            if (firstBrace >= 0 && lastBrace > firstBrace) {
                candidates.push(raw.slice(firstBrace, lastBrace + 1));
            }

            for (const candidate of candidates) {
                try {
                    const parsed = JSON.parse(candidate);
                    if (parsed && typeof parsed === 'object' && !Array.isArray(parsed)) {
                        return parsed;
                    }
                } catch {
                    continue;
                }
            }

            return null;
        }

        async function polishTranscriptCaptions(captions) {
            if (!Array.isArray(captions) || captions.length === 0) {
                return captions;
            }

            const totalChars = captions.reduce((sum, caption) => sum + String(caption.text || '').length, 0);
            if (captions.length > 800 || totalChars > 80000) {
                console.log('[Meet] Transcript too large for polishing, using cleaned version.');
                return captions;
            }

            const turnsForModel = captions.map((caption, index) => ({
                sourceIndex: index + 1,
                speaker: caption.speaker,
                time: formatCaptionTime(caption.timestamp),
                text: caption.text
            }));

            const polishPrompt = `Fix this ASR transcript. Return ONLY valid JSON.

Input: ${JSON.stringify(turnsForModel)}

Output shape: { "turns": [{ "sourceIndex": 1, "speaker": "...", "text": "..." }] }

Rules:
- Fix misheard words by context (e.g. "Mcptation" → "MCP integration", "rubra tap" → "Rubrik app")
- Capitalize technical terms correctly: OAuth, API, GitHub, MCP, OpenAI, GPT, etc.
- Fix grammar and sentence flow naturally
- Remove ASR filler noise and pure-noise turns
- Split multi-speaker turns; keep original speaker names unless clearly wrong
- Do not add speaker names inside text field
- Do not invent facts or add information not present
- No markdown, no code fences, no extra keys`;

            try {
                console.log('[Meet] Polishing transcript text with OpenAI...');

                const polishCompletion = await openai.chat.completions.create({
                    model: OPENAI_MODEL,
                    messages: [
                        {
                            role: 'system',
                            content: 'You fix ASR transcript errors. Return JSON only.'
                        },
                        {
                            role: 'user',
                            content: polishPrompt
                        }
                    ],
                    temperature: 0.1,
                    max_tokens: 16000
                });

                const raw = polishCompletion.choices?.[0]?.message?.content || '';
                const parsed = parseJsonFromModelOutput(raw);
                const modelTurns = Array.isArray(parsed?.turns) ? parsed.turns : [];
                if (modelTurns.length === 0) {
                    return captions;
                }

                const rebuiltTurns = [];
                for (const turn of modelTurns) {
                    const sourceIndex = Number(turn?.sourceIndex ?? turn?.index);
                    if (!Number.isInteger(sourceIndex) || sourceIndex < 1 || sourceIndex > captions.length) {
                        continue;
                    }

                    const sourceCaption = captions[sourceIndex - 1];
                    const speaker = sanitizeSpeakerName(turn?.speaker, sourceCaption?.speaker || 'Unknown Speaker');
                    const text = sanitizeTranscriptText(turn?.text);

                    if (!text) {
                        continue;
                    }

                    rebuiltTurns.push({
                        speaker,
                        text,
                        timestamp: sourceCaption?.timestamp || null
                    });
                }

                if (rebuiltTurns.length === 0) {
                    return captions;
                }

                return deduplicateCaptions(rebuiltTurns);
            } catch (error) {
                console.error('[Meet] Transcript polishing failed, using cleaned transcript:', error.message);
                return captions;
            }
        }

        const normalizedCaptions = normalizeCaptions(session.captions);
        const cleanedCaptions = deduplicateCaptions(normalizedCaptions);
        const polishedCaptions = await polishTranscriptCaptions(cleanedCaptions);
        const transcriptStepByStep = formatStepByStepTranscript(polishedCaptions);
        const transcriptSpeakerWise = formatSpeakerWiseTranscript(polishedCaptions);

        if (!transcriptStepByStep || transcriptStepByStep.trim().length === 0) {
            return res.status(400).json({ error: 'No captions captured' });
        }

        // Generate AI summary using OpenAI
        const summaryPrompt = `Meeting: ${session.metadata.title || 'Google Meet'} (${new Date(session.startTime).toLocaleString()})

Transcript:
${transcriptStepByStep}

Generate meeting notes in this exact structure:

## Key Discussion Points
## Decisions Made
## Action Items (Owner | Action | Due)
## Next Steps
## Open Questions

Rules: Concise factual points only. If a section has nothing, write "- None identified." No intro/closing text.`;

        console.log('[Meet] Generating summary with OpenAI...');

        const completion = await openai.chat.completions.create({
            model: OPENAI_MODEL,
            messages: [
                {
                    role: 'system',
                    content: 'You are a meeting-notes formatter. Follow the requested structure exactly.'
                },
                {
                    role: 'user',
                    content: summaryPrompt
                }
            ],
            temperature: 0.2,
            max_tokens: 4000
        });

        const summary = completion.choices[0].message.content;

        // Create Google Doc with summary + transcript
        const docTitle = `${session.metadata.title || 'Meeting'} - Notes (${new Date().toLocaleDateString()})`;

        const docContent = `# ${session.metadata.title || 'Meeting Notes'}

**Date:** ${new Date(session.startTime).toLocaleString()}
**Meeting Code:** ${session.metadata.meetingCode || 'N/A'}
**Duration:** ${Math.round((new Date() - new Date(session.startTime)) / 60000)} minutes
**Attendees:** ${session.metadata.attendees?.join(', ') || 'N/A'}

---

## AI Summary

${summary}

---

## Full Transcript (Polished Step-by-Step)

${transcriptStepByStep}

---

## Full Transcript (Polished Speaker-Wise)

${transcriptSpeakerWise}

---

*Generated by Google Meet Note-Taker*
`;

        console.log('[Meet] Creating Google Doc...');

        const createResult = await createDocument({
            title: docTitle,
            content: docContent
        });

        if (!createResult || !createResult.documentId) {
            throw new Error('Failed to create Google Doc: No document ID returned');
        }

        const documentId = createResult.documentId;
        const docUrl = `https://docs.google.com/document/d/${documentId}/edit`;

        console.log('[Meet] Google Doc created:', docUrl);

        // Clean up session
        const attendeeEmails = Array.isArray(session.metadata.attendees)
            ? session.metadata.attendees
                .map(email => String(email || '').trim())
                .filter(Boolean)
            : [];

        meetSessions.delete(sessionId);

        res.json({
            success: true,
            documentId,
            docUrl,
            summary,
            captionCount: polishedCaptions.length,
            attendees: attendeeEmails,
            autoShared: false
        });

    } catch (error) {
        console.error('Error finalizing meeting:', error);
        res.status(500).json({ error: 'Internal server error' });
    }
});

// Share Google Doc with attendees
app.post('/api/meet/share', async (req, res) => {
    try {
        const { documentId, emails } = req.body;

        if (!documentId) {
            return res.status(400).json({ error: 'Document ID is required' });
        }

        if (!emails || !Array.isArray(emails) || emails.length === 0) {
            return res.status(400).json({ error: 'At least one email address is required' });
        }

        if (!driveClient || !hasDriveScope()) {
            return res.status(503).json({
                error: 'Google Drive service not available',
                setupRequired: true
            });
        }

        console.log(`[Meet] Sharing doc ${documentId} with ${emails.length} attendee(s)`);

        const results = [];
        const errors = [];

        for (const email of emails) {
            const trimmedEmail = email.trim();
            if (!trimmedEmail || !trimmedEmail.includes('@')) {
                errors.push({ email: trimmedEmail, error: 'Invalid email format' });
                continue;
            }

            try {
                await shareDriveFile({
                    fileId: documentId,
                    emailAddress: trimmedEmail,
                    role: 'writer',
                    sendNotificationEmail: true
                });
                results.push({ email: trimmedEmail, success: true });
                console.log(`[Meet] Successfully shared with ${trimmedEmail}`);
            } catch (shareError) {
                errors.push({ email: trimmedEmail, error: shareError.message });
                console.error(`[Meet] Failed to share with ${trimmedEmail}:`, shareError.message);
            }
        }

        res.json({
            success: true,
            shared: results.length,
            failed: errors.length,
            results,
            errors
        });

    } catch (error) {
        console.error('Error sharing document:', error);
        res.status(500).json({ error: 'Internal server error' });
    }
});

// Non-streaming chat endpoint (legacy, for backward compatibility)
app.post('/api/chat', async (req, res) => {
    const { message, history = [], attachedFiles = [] } = req.body;
    try {
        if (typeof message !== 'string' || !message.trim()) {
            return res.status(400).json({ error: 'message must be a non-empty string' });
        }
        if (message.length > MAX_CHAT_MESSAGE_CHARS) {
            return res.status(400).json({ error: `message exceeds max length (${MAX_CHAT_MESSAGE_CHARS} chars)` });
        }
        if (Array.isArray(history) && history.length > MAX_CHAT_HISTORY_ITEMS) {
            return res.status(400).json({ error: `history exceeds max items (${MAX_CHAT_HISTORY_ITEMS})` });
        }
        if (Array.isArray(attachedFiles) && attachedFiles.length > MAX_CHAT_ATTACHED_FILES) {
            return res.status(400).json({ error: `too many attached files (max ${MAX_CHAT_ATTACHED_FILES})` });
        }

        // Process attached files
        const fileContexts = [];
        if (Array.isArray(attachedFiles) && attachedFiles.length > 0) {
            for (const fileId of attachedFiles) {
                const fileInfo = uploadedFiles.get(fileId);
                if (fileInfo) {
                    fileContexts.push({
                        fileId,
                        name: fileInfo.originalName,
                        size: fileInfo.size,
                        mimeType: fileInfo.mimeType,
                        localPath: fileInfo.path
                    });
                }
            }
        }

        const safeHistory = buildSafeChatHistory(history);
        const result = await runAgentConversation({ message, history: safeHistory, attachedFiles: fileContexts });
        res.json(result);
    } catch (error) {
        console.error('Chat error:', error);
        res.status(500).json({ error: 'Internal server error' });
    }
});

// ============================================================
//  START SERVER
// ============================================================
const credentialsConfigured = initOAuthClient();
initGitHubClient();
initSheetsMcpClient().catch(error => {
    console.error('Sheets MCP bootstrap failed:', error?.message || error);
});
initOutlookClient().catch(error => {
    console.error('Outlook init failed:', error?.message || error);
});
initGcsClient();
loadScheduledTasksFromDisk();
startScheduledTaskRunner();

const totalTools = gmailTools.length + calendarTools.length + gchatTools.length + driveTools.length + sheetsTools.length + sheetsMcpTools.length + githubTools.length + outlookTools.length + docsTools.length + meetingTranscriptionTools.length + teamsTools.length + gcsTools.length;
let startupErrorHandled = false;
const handleStartupError = (error, source) => {
    if (startupErrorHandled) return;
    startupErrorHandled = true;

    const errorCode = error?.code || 'UNKNOWN';
    const errorMessage = error?.message || 'Unknown startup error';

    if (errorCode === 'EADDRINUSE') {
        console.error(`Port ${PORT} is already in use (${source}).`);
        console.error('Stop the existing process on this port or set a different PORT in .env.');
        process.exit(1);
        return;
    }

    console.error(`Server startup error [${source}] (${errorCode}): ${errorMessage}`);
    process.exit(1);
};

const httpServer = app.listen(PORT, HOST, () => {
    console.log(`\nAI Agent Server running at http://${HOST}:${PORT}`);
    console.log(`OpenAI model: ${OPENAI_MODEL}${OPENAI_FALLBACK_MODEL ? ` (fallback: ${OPENAI_FALLBACK_MODEL})` : ''}, retries: ${OPENAI_CHAT_MAX_RETRIES}, temperature: ${OPENAI_TEMPERATURE}${OPENAI_MAX_OUTPUT_TOKENS ? `, max output tokens: ${OPENAI_MAX_OUTPUT_TOKENS}` : ''}`);
    console.log(`Total tools available: ${totalTools} (Gmail: ${gmailTools.length}, Calendar: ${calendarTools.length}, Chat: ${gchatTools.length}, Drive: ${driveTools.length}, Sheets: ${sheetsTools.length}, Sheets MCP: ${sheetsMcpTools.length}, Docs: ${docsTools.length}, Meeting Transcription: ${meetingTranscriptionTools.length}, GitHub: ${githubTools.length}, Outlook: ${outlookTools.length}, Teams: ${teamsTools.length}, GCS: ${gcsTools.length})`);
    console.log(`Gmail: ${gmailClient ? 'Connected' : 'Not connected'}`);
    console.log(`Calendar: ${calendarClient && hasCalendarScope() ? 'Connected' : 'Not connected'}`);
    console.log(`Google Chat: ${gchatClient && hasGchatScopes() ? 'Connected' : 'Not connected'}`);
    console.log(`Google Drive: ${driveClient && hasDriveScope() ? 'Connected' : 'Not connected'}`);
    console.log(`Google Sheets: ${sheetsClient && hasSheetsScope() ? 'Connected' : 'Not connected'}`);
    console.log(`Google Sheets MCP: ${sheetsMcpClient ? `Connected (${sheetsMcpTools.length} tools)` : `Not connected${sheetsMcpError ? ` (${sheetsMcpError})` : ''}`}`);
    console.log(`Google Docs: ${docsClient && hasDocsScope() ? 'Connected' : 'Not connected'}`);
    console.log(`Meeting Transcription: ${driveClient && hasDriveScope() ? 'Connected' : 'Not connected'}`);
    console.log(`GitHub: ${octokitClient ? 'Connected' : 'Not connected'}`);
    console.log(`Outlook: ${outlookAccessToken ? `Connected (${outlookUserEmail || 'unknown'})` : 'Not connected'}`);
    console.log(`Microsoft Teams: ${outlookAccessToken && hasTeamsScopes() ? 'Connected' : 'Not connected'}`);
    console.log(`GCS: ${gcsAuthenticated ? `Connected (project: ${gcsProjectId})` : 'Not connected'}`);
    console.log(`Timer Tasks: ${scheduledTasks.length} configured (${scheduledTasks.filter(task => task.enabled).length} enabled)`);
});
httpServer.once('error', (error) => handleStartupError(error, 'http'));

// WebSocket Server for Google Meet Note-Taker
const wss = new WebSocket.Server({ server: httpServer, path: '/meet-notes' });
wss.once('error', (error) => handleStartupError(error, 'websocket'));
const meetSessions = new Map(); // sessionId -> { metadata, captions[], startTime, ws }
const MAX_MEET_SESSIONS = Math.max(1, Number.parseInt(process.env.MAX_MEET_SESSIONS || '30', 10) || 30);
const MAX_CAPTIONS_PER_SESSION = Math.max(100, Number.parseInt(process.env.MAX_CAPTIONS_PER_SESSION || '20000', 10) || 20000);
const MAX_CAPTION_TEXT_CHARS = Math.max(100, Number.parseInt(process.env.MAX_CAPTION_TEXT_CHARS || '2000', 10) || 2000);
const MAX_SESSION_ID_CHARS = 120;
const TRUSTED_WS_ORIGIN_PATTERNS = [
    /^http:\/\/localhost:\d+$/i,
    /^http:\/\/127\.0\.0\.1:\d+$/i,
    /^chrome-extension:\/\/[a-p]{32}$/i
];

function isTrustedWsOrigin(origin) {
    const raw = String(origin || '').trim();
    if (!raw) return false;
    return TRUSTED_WS_ORIGIN_PATTERNS.some(pattern => pattern.test(raw));
}

function sanitizeMeetMetadata(raw) {
    const source = (raw && typeof raw === 'object' && !Array.isArray(raw)) ? raw : {};
    const attendees = Array.isArray(source.attendees)
        ? source.attendees
            .map(value => String(value || '').trim())
            .filter(Boolean)
            .slice(0, 100)
        : [];

    return {
        title: String(source.title || '').trim().slice(0, 200) || 'Google Meet',
        meetingCode: String(source.meetingCode || '').trim().slice(0, 120),
        startTime: source.startTime || null,
        endTime: source.endTime || null,
        eventId: source.eventId || null,
        hangoutLink: source.hangoutLink || null,
        attendees
    };
}

function sanitizeCaptionPayload(raw) {
    if (!raw || typeof raw !== 'object') return null;
    const speaker = String(raw.speaker || '').replace(/\s+/g, ' ').trim().slice(0, 120) || 'Unknown Speaker';
    const text = String(raw.text || '').replace(/\s+/g, ' ').trim().slice(0, MAX_CAPTION_TEXT_CHARS);
    if (!text) return null;

    const parsed = raw.timestamp ? new Date(raw.timestamp) : null;
    const timestamp = parsed && !Number.isNaN(parsed.getTime())
        ? parsed.toISOString()
        : new Date().toISOString();

    return { speaker, text, timestamp };
}

function isValidSessionId(sessionId) {
    const id = String(sessionId || '').trim();
    if (!id || id.length > MAX_SESSION_ID_CHARS) return false;
    return /^[A-Za-z0-9._:-]+$/.test(id);
}

wss.on('connection', (ws, request) => {
    const origin = request?.headers?.origin;
    const remoteAddress = request?.socket?.remoteAddress || '';
    const trustedOrigin = isTrustedWsOrigin(origin);
    const localPeer = isLoopbackAddress(remoteAddress);

    if (!trustedOrigin && !(localPeer && !origin)) {
        console.warn(`[WebSocket] Rejected untrusted connection. origin=${origin || '(missing)'} ip=${remoteAddress || '(unknown)'}`);
        ws.close(1008, 'Untrusted origin');
        return;
    }

    console.log('[WebSocket] Client connected');

    ws.on('message', (message) => {
        try {
            const data = JSON.parse(message.toString());
            const messageType = String(data?.type || '').trim();
            const sessionId = String(data?.sessionId || '').trim();
            console.log('[WebSocket] Message received:', messageType || '(unknown)');

            if (!isValidSessionId(sessionId)) {
                ws.send(JSON.stringify({ type: 'error', message: 'Invalid or missing sessionId' }));
                return;
            }

            if (messageType === 'session_start') {
                const existingSession = meetSessions.get(sessionId);
                if (!existingSession && meetSessions.size >= MAX_MEET_SESSIONS) {
                    ws.send(JSON.stringify({ type: 'error', message: 'Too many active sessions' }));
                    return;
                }

                if (existingSession) {
                    existingSession.metadata = sanitizeMeetMetadata(data.metadata);
                    existingSession.ws = ws;
                } else {
                    meetSessions.set(sessionId, {
                        metadata: sanitizeMeetMetadata(data.metadata),
                        captions: [],
                        startTime: new Date(),
                        ws
                    });
                }

                console.log(`[WebSocket] Session started: ${sessionId}${existingSession ? ' (resumed)' : ''}`);
                ws.send(JSON.stringify({ type: 'session_created', sessionId }));
                return;
            }

            const session = meetSessions.get(sessionId);
            if (!session) {
                ws.send(JSON.stringify({ type: 'error', message: 'Session not found' }));
                return;
            }

            if (session.ws !== ws) {
                ws.send(JSON.stringify({ type: 'error', message: 'Session access denied' }));
                return;
            }

            if (messageType === 'caption') {
                if (session.captions.length >= MAX_CAPTIONS_PER_SESSION) {
                    ws.send(JSON.stringify({ type: 'error', message: 'Caption limit reached for this session' }));
                    return;
                }

                const caption = sanitizeCaptionPayload(data.caption);
                if (!caption) {
                    ws.send(JSON.stringify({ type: 'error', message: 'Invalid caption payload' }));
                    return;
                }

                session.captions.push(caption);

                // Broadcast to all connected clients (for live view)
                wss.clients.forEach(client => {
                    if (client.readyState === WebSocket.OPEN) {
                        client.send(JSON.stringify({
                            type: 'caption_added',
                            caption,
                            sessionId
                        }));
                    }
                });
                return;
            }

            if (messageType === 'session_end') {
                console.log(`[WebSocket] Session ending: ${sessionId}`);
                ws.send(JSON.stringify({ type: 'session_ended', sessionId }));
                return;
            }

            ws.send(JSON.stringify({ type: 'error', message: `Unsupported message type: ${messageType || 'unknown'}` }));
        } catch (error) {
            console.error('[WebSocket] Error processing message:', error);
            ws.send(JSON.stringify({
                type: 'error',
                message: error.message
            }));
        }
    });

    ws.on('close', () => {
        console.log('[WebSocket] Client disconnected');
    });

    ws.on('error', (error) => {
        console.error('[WebSocket] Error:', error);
    });
});

console.log('WebSocket server initialized on path: /meet-notes');

process.on('SIGINT', async () => {
    if (schedulerInterval) clearInterval(schedulerInterval);
    await closeSheetsMcpClient();
    process.exit(0);
});

process.on('SIGTERM', async () => {
    if (schedulerInterval) clearInterval(schedulerInterval);
    await closeSheetsMcpClient();
    process.exit(0);
});




