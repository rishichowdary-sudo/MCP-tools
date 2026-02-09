require('dotenv').config();
const express = require('express');
const cors = require('cors');
const { google } = require('googleapis');
const OpenAI = require('openai');
const { Octokit } = require('octokit');
const { Client: McpClient } = require('@modelcontextprotocol/sdk/client');
const { StdioClientTransport } = require('@modelcontextprotocol/sdk/client/stdio.js');
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors());
app.use(express.json());
app.use(express.static('public'));

// OpenAI setup
const openai = new OpenAI({
    apiKey: process.env.OPENAI_API_KEY
});

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
                    if (requiresStrictHeaderMatch && !matchedIdentity) continue;
                    const emails = extractEmailsFromText(value);

                    for (const email of emails) {
                        if (ownEmail && email === ownEmail) continue;
                        if (isDisallowedAttendeeEmail(email)) continue;
                        const localPart = normalizeIdentity(email.split('@')[0] || '');
                        const matchedLocalPart = keyTokens.length > 1
                            ? keyTokens.every(token => localPart.includes(token))
                            : localPart.includes(key);
                        if (requiresStrictHeaderMatch && !matchedIdentity && !matchedLocalPart) continue;
                        const existing = scoreByEmail.get(email) || 0;
                        const bump = matchedIdentity ? 5 : (matchedLocalPart ? 4 : 1);
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
    return ranked[0][0];
}

async function normalizeEventAttendees(attendeesInput = []) {
    if (!Array.isArray(attendeesInput) || attendeesInput.length === 0) return [];

    const resolved = [];
    const unresolved = [];

    for (const item of attendeesInput) {
        const raw = String(item || '').trim();
        if (!raw) continue;

        if (isValidEmail(raw) && !isDisallowedAttendeeEmail(raw)) {
            resolved.push(raw.toLowerCase());
            continue;
        }

        const found = await resolveEmailFromGmailHistory(raw);
        if (found) resolved.push(found.toLowerCase());
        else unresolved.push(raw);
    }

    const deduped = [...new Set(resolved)];
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
    const event = { summary };
    if (description) event.description = description;
    if (location) event.location = location;
    if (startDateTime) event.start = { dateTime: startDateTime, timeZone: timeZone || 'UTC' };
    else if (startDate) event.start = { date: startDate };
    if (endDateTime) event.end = { dateTime: endDateTime, timeZone: timeZone || 'UTC' };
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
async function createMeetEvent({ calendarId = 'primary', summary, description, startDateTime, endDateTime, attendees = [], timeZone = 'UTC' }) {
    if (!calendarClient) throw new Error('Calendar not authenticated');
    if (!summary || !startDateTime || !endDateTime) {
        throw new Error('summary, startDateTime, and endDateTime are required to create a Meet event');
    }

    const resolvedAttendees = await normalizeEventAttendees(attendees);
    const requestBody = {
        summary,
        description,
        start: { dateTime: startDateTime, timeZone },
        end: { dateTime: endDateTime, timeZone },
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
        const resolvedAdditions = await normalizeEventAttendees(addAttendees);
        const existingEmails = new Set(attendees.map(a => a.email));
        for (const email of resolvedAdditions) {
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
//  GOOGLE CHAT TOOL IMPLEMENTATIONS
// ============================================================

function normalizeSpaceName(spaceId) {
    if (!spaceId) return '';
    return spaceId.startsWith('spaces/') ? spaceId : `spaces/${spaceId}`;
}

async function listChatSpaces({ maxResults = 20 }) {
    if (!gchatClient) throw new Error('Google Chat not authenticated');
    const response = await gchatClient.spaces.list({ pageSize: maxResults });
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
    const response = await gchatClient.spaces.messages.list({
        parent,
        pageSize: maxResults
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

async function listDriveFiles({ query, pageSize = 100, orderBy = 'modifiedTime desc', includeTrashed = false }) {
    if (!driveClient) throw new Error('Google Drive not authenticated');
    const qParts = [];
    if (!includeTrashed) qParts.push('trashed = false');
    const queryText = String(query || '').trim();
    const primaryClause = buildDriveQueryClause(queryText);
    if (primaryClause) qParts.push(primaryClause);

    let response;
    try {
        response = await driveClient.files.list({
            q: qParts.length > 0 ? qParts.join(' and ') : undefined,
            pageSize,
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
            pageSize,
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

async function createDriveFile({ name, content = '', mimeType = 'text/plain', parentId }) {
    if (!driveClient) throw new Error('Google Drive not authenticated');
    if (!name) throw new Error('name is required');

    const requestBody = { name };
    if (parentId) requestBody.parents = [parentId];

    const response = await driveClient.files.create({
        requestBody,
        media: { mimeType, body: content },
        fields: 'id,name,mimeType,webViewLink,parents,size,modifiedTime'
    });
    return {
        success: true,
        file: normalizeDriveFile(response.data),
        message: `File "${response.data.name}" created`
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

async function downloadDriveFile({ fileId, maxBytes = 200000 }) {
    if (!driveClient) throw new Error('Google Drive not authenticated');
    if (!fileId) throw new Error('fileId is required');

    const meta = await driveClient.files.get({
        fileId,
        fields: 'id,name,mimeType,size'
    });

    const mimeType = meta.data.mimeType || '';
    if (mimeType === 'application/vnd.google-apps.spreadsheet') {
        throw new Error('This file is a Google Sheet. Use Sheets tools (read_sheet_values, get_spreadsheet) to access data.');
    }

    let rawBytes;
    if (mimeType.startsWith('application/vnd.google-apps.')) {
        const exportResponse = await driveClient.files.export(
            { fileId, mimeType: 'text/plain' },
            { responseType: 'arraybuffer' }
        );
        rawBytes = Buffer.from(exportResponse.data);
    } else {
        const fileResponse = await driveClient.files.get(
            { fileId, alt: 'media' },
            { responseType: 'arraybuffer' }
        );
        rawBytes = Buffer.from(fileResponse.data);
    }

    const limit = Math.max(1024, Number(maxBytes) || 200000);
    const truncated = rawBytes.length > limit;
    const body = truncated ? rawBytes.subarray(0, limit) : rawBytes;
    const content = body.toString('utf8');

    return {
        fileId,
        name: meta.data.name,
        mimeType,
        content,
        byteLength: rawBytes.length,
        returnedBytes: body.length,
        truncated,
        message: truncated
            ? `Downloaded truncated content for "${meta.data.name}" (${body.length}/${rawBytes.length} bytes)`
            : `Downloaded file "${meta.data.name}"`
    };
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
            pageSize: maxResults,
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
            pageSize: maxResults,
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
            pageSize: Math.max(maxResults, 200),
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
    const qParts = ["mimeType='application/vnd.google-apps.document'", 'trashed = false'];
    if (query) qParts.push(`name contains '${query.replace(/'/g, "\\'")}'`);
    const response = await driveClient.files.list({
        q: qParts.join(' and '),
        pageSize,
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
    const toRecipients = (Array.isArray(to) ? to : [to]).map(addr => ({
        emailAddress: { address: addr }
    }));
    const message = {
        subject: subject || '',
        body: { contentType: 'HTML', content: body || '' },
        toRecipients
    };
    if (cc) {
        message.ccRecipients = (Array.isArray(cc) ? cc : [cc]).map(addr => ({
            emailAddress: { address: addr }
        }));
    }
    if (bcc) {
        message.bccRecipients = (Array.isArray(bcc) ? bcc : [bcc]).map(addr => ({
            emailAddress: { address: addr }
        }));
    }
    await outlookGraphFetch('/me/sendMail', {
        method: 'POST',
        body: JSON.stringify({ message, saveToSentItems: true })
    });
    return { success: true, message: `Email sent to ${Array.isArray(to) ? to.join(', ') : to}` };
}

// 2. List Emails
async function outlookListEmails({ maxResults = 20, folder = 'inbox' }) {
    const data = await outlookGraphFetch(
        `/me/mailFolders/${folder}/messages?$top=${maxResults}&$select=id,subject,from,receivedDateTime,isRead,bodyPreview,hasAttachments&$orderby=receivedDateTime desc`
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
    const data = await outlookGraphFetch(
        `/me/messages?$search="${encodeURIComponent(query)}"&$top=${maxResults}&$select=id,subject,from,receivedDateTime,isRead,bodyPreview,hasAttachments`
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
    const toRecipients = (Array.isArray(to) ? to : [to]).map(addr => ({
        emailAddress: { address: addr }
    }));
    await outlookGraphFetch(`/me/messages/${messageId}/forward`, {
        method: 'POST',
        body: JSON.stringify({ comment: comment || '', toRecipients })
    });
    return { success: true, message: `Email forwarded to ${Array.isArray(to) ? to.join(', ') : to}` };
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
    const message = {
        subject: subject || '',
        body: { contentType: 'HTML', content: body || '' }
    };
    if (to) {
        message.toRecipients = (Array.isArray(to) ? to : [to]).map(addr => ({
            emailAddress: { address: addr }
        }));
    }
    if (cc) {
        message.ccRecipients = (Array.isArray(cc) ? cc : [cc]).map(addr => ({
            emailAddress: { address: addr }
        }));
    }
    if (bcc) {
        message.bccRecipients = (Array.isArray(bcc) ? bcc : [bcc]).map(addr => ({
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
    const data = await outlookGraphFetch(
        `/me/mailFolders/drafts/messages?$top=${maxResults}&$select=id,subject,toRecipients,createdDateTime,bodyPreview&$orderby=createdDateTime desc`
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
    const data = await outlookGraphFetch(`/teams/${teamId}/channels/${channelId}/messages?$top=${top}`);
    const messages = (data.value || []).map(m => ({
        id: m.id, from: m.from?.user?.displayName || 'Unknown',
        body: m.body?.content?.slice(0, 500), contentType: m.body?.contentType,
        createdDateTime: m.createdDateTime
    }));
    return { messages, count: messages.length };
}

async function teamsListChats({ top = 20 }) {
    const data = await outlookGraphFetch(`/me/chats?$top=${top}&$expand=members`);
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
    const data = await outlookGraphFetch(`/chats/${chatId}/messages?$top=${top}`);
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
//  TOOL DEFINITIONS FOR OPENAI
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
    { type: "function", function: { name: "create_event", description: "Create a new calendar event with optional attendees, location, recurrence, and optional Google Meet link.", parameters: { type: "object", properties: { calendarId: { type: "string", description: "Calendar ID (default: primary)" }, summary: { type: "string", description: "Event title" }, description: { type: "string", description: "Event description" }, location: { type: "string", description: "Event location" }, startDateTime: { type: "string", description: "Start datetime in ISO 8601 (for timed events)" }, endDateTime: { type: "string", description: "End datetime in ISO 8601 (for timed events)" }, startDate: { type: "string", description: "Start date YYYY-MM-DD (for all-day events)" }, endDate: { type: "string", description: "End date YYYY-MM-DD (for all-day events)" }, attendees: { type: "array", items: { type: "string" }, description: "Attendee email addresses" }, recurrence: { type: "array", items: { type: "string" }, description: "RRULE strings, e.g. ['RRULE:FREQ=WEEKLY;COUNT=5']" }, timeZone: { type: "string", description: "Time zone (default: UTC)" }, createMeetLink: { type: "boolean", description: "If true, create a Google Meet link for this event" } }, required: ["summary"] } } },
    { type: "function", function: { name: "create_meet_event", description: "Create a Google Calendar event with a Google Meet link.", parameters: { type: "object", properties: { calendarId: { type: "string", description: "Calendar ID (default: primary)" }, summary: { type: "string", description: "Meeting title" }, description: { type: "string", description: "Meeting description" }, startDateTime: { type: "string", description: "Start datetime in ISO 8601" }, endDateTime: { type: "string", description: "End datetime in ISO 8601" }, attendees: { type: "array", items: { type: "string" }, description: "Attendee email addresses" }, timeZone: { type: "string", description: "Time zone (default: UTC)" } }, required: ["summary", "startDateTime", "endDateTime"] } } },
    { type: "function", function: { name: "add_meet_link_to_event", description: "Add a Google Meet link to an existing calendar event.", parameters: { type: "object", properties: { calendarId: { type: "string", description: "Calendar ID (default: primary)" }, eventId: { type: "string", description: "The event ID" } }, required: ["eventId"] } } },
    { type: "function", function: { name: "update_event", description: "Update an existing calendar event's title, time, location, or description.", parameters: { type: "object", properties: { calendarId: { type: "string", description: "Calendar ID (default: primary)" }, eventId: { type: "string", description: "The event ID to update" }, summary: { type: "string", description: "New event title" }, description: { type: "string", description: "New description" }, location: { type: "string", description: "New location" }, startDateTime: { type: "string", description: "New start datetime" }, endDateTime: { type: "string", description: "New end datetime" }, startDate: { type: "string", description: "New start date (all-day)" }, endDate: { type: "string", description: "New end date (all-day)" }, timeZone: { type: "string", description: "Time zone" } }, required: ["eventId"] } } },
    { type: "function", function: { name: "delete_event", description: "Delete a calendar event.", parameters: { type: "object", properties: { calendarId: { type: "string", description: "Calendar ID (default: primary)" }, eventId: { type: "string", description: "The event ID to delete" } }, required: ["eventId"] } } },
    { type: "function", function: { name: "list_calendars", description: "List all calendars accessible to the user.", parameters: { type: "object", properties: {} } } },
    { type: "function", function: { name: "create_calendar", description: "Create a new calendar.", parameters: { type: "object", properties: { summary: { type: "string", description: "Calendar name" }, description: { type: "string", description: "Calendar description" }, timeZone: { type: "string", description: "Time zone" } }, required: ["summary"] } } },
    { type: "function", function: { name: "quick_add_event", description: "Quickly create an event using natural language (e.g. 'Meeting tomorrow at 3pm').", parameters: { type: "object", properties: { calendarId: { type: "string", description: "Calendar ID (default: primary)" }, text: { type: "string", description: "Natural language event description" } }, required: ["text"] } } },
    { type: "function", function: { name: "get_free_busy", description: "Check free/busy status for specified calendar IDs in a time range. If calendarIds is omitted, it checks only your primary calendar.", parameters: { type: "object", properties: { timeMin: { type: "string", description: "Start of time range (ISO 8601)" }, timeMax: { type: "string", description: "End of time range (ISO 8601)" }, calendarIds: { type: "array", items: { type: "string" }, description: "Calendar IDs to check (default: ['primary'])" } }, required: ["timeMin", "timeMax"] } } },
    { type: "function", function: { name: "check_person_availability", description: "Check one person's calendar availability and return free slots. Requires access to that person's calendar free/busy data.", parameters: { type: "object", properties: { person: { type: "string", description: "Person name or email to resolve from Gmail history" }, email: { type: "string", description: "Exact person email (preferred if known)" }, calendarId: { type: "string", description: "Calendar ID override (if known)" }, timeMin: { type: "string", description: "Start of time range (ISO 8601)" }, timeMax: { type: "string", description: "End of time range (ISO 8601)" }, durationMinutes: { type: "integer", description: "Minimum free slot length in minutes (default 30)" } }, required: ["timeMin", "timeMax"] } } },
    { type: "function", function: { name: "find_common_free_slots", description: "Find common free slots across multiple people and/or calendars.", parameters: { type: "object", properties: { timeMin: { type: "string", description: "Start of time range (ISO 8601)" }, timeMax: { type: "string", description: "End of time range (ISO 8601)" }, people: { type: "array", items: { type: "string" }, description: "Names or emails to resolve and include" }, calendarIds: { type: "array", items: { type: "string" }, description: "Optional calendar IDs to include" }, includePrimary: { type: "boolean", description: "Include your primary calendar (default true)" }, durationMinutes: { type: "integer", description: "Minimum free slot length in minutes (default 30)" } }, required: ["timeMin", "timeMax"] } } },
    { type: "function", function: { name: "list_recurring_instances", description: "List individual occurrences of a recurring event.", parameters: { type: "object", properties: { calendarId: { type: "string", description: "Calendar ID (default: primary)" }, eventId: { type: "string", description: "The recurring event ID" }, maxResults: { type: "integer", description: "Max instances to return" }, timeMin: { type: "string", description: "Start time filter" }, timeMax: { type: "string", description: "End time filter" } }, required: ["eventId"] } } },
    { type: "function", function: { name: "move_event", description: "Move an event from one calendar to another.", parameters: { type: "object", properties: { calendarId: { type: "string", description: "Source calendar ID (default: primary)" }, eventId: { type: "string", description: "The event ID to move" }, destinationCalendarId: { type: "string", description: "Destination calendar ID" } }, required: ["eventId", "destinationCalendarId"] } } },
    { type: "function", function: { name: "update_event_attendees", description: "Add or remove attendees from a calendar event.", parameters: { type: "object", properties: { calendarId: { type: "string", description: "Calendar ID (default: primary)" }, eventId: { type: "string", description: "The event ID" }, addAttendees: { type: "array", items: { type: "string" }, description: "Email addresses to add" }, removeAttendees: { type: "array", items: { type: "string" }, description: "Email addresses to remove" } }, required: ["eventId"] } } },
    { type: "function", function: { name: "get_calendar_colors", description: "Get available color options for calendars and events.", parameters: { type: "object", properties: {} } } },
    { type: "function", function: { name: "clear_calendar", description: "Clear all events from a calendar. WARNING: This is destructive!", parameters: { type: "object", properties: { calendarId: { type: "string", description: "The calendar ID to clear (cannot be primary)" } }, required: ["calendarId"] } } },
    { type: "function", function: { name: "watch_events", description: "Set up push notifications for calendar changes (requires a public webhook URL).", parameters: { type: "object", properties: { calendarId: { type: "string", description: "Calendar ID (default: primary)" }, webhookUrl: { type: "string", description: "Public webhook URL to receive notifications" } }, required: ["webhookUrl"] } } }
];

const gchatTools = [
    { type: "function", function: { name: "list_chat_spaces", description: "List Google Chat spaces available to the authenticated user.", parameters: { type: "object", properties: { maxResults: { type: "integer", description: "Max spaces to return (default 20)" } } } } },
    { type: "function", function: { name: "send_chat_message", description: "Send a text message to a Google Chat space.", parameters: { type: "object", properties: { spaceId: { type: "string", description: "Space ID or full name like spaces/AAAA..." }, text: { type: "string", description: "Message text to send" } }, required: ["spaceId", "text"] } } },
    { type: "function", function: { name: "list_chat_messages", description: "List recent messages in a Google Chat space.", parameters: { type: "object", properties: { spaceId: { type: "string", description: "Space ID or full name like spaces/AAAA..." }, maxResults: { type: "integer", description: "Max messages to return (default 20)" } }, required: ["spaceId"] } } }
];

const driveTools = [
    { type: "function", function: { name: "list_drive_files", description: "List Google Drive files/folders you can access. Supports Drive query syntax.", parameters: { type: "object", properties: { query: { type: "string", description: "Drive query string, e.g. name contains 'Q1' and mimeType contains 'spreadsheet'" }, pageSize: { type: "integer", description: "Max files to return (default 25)" }, orderBy: { type: "string", description: "Sort order (default 'modifiedTime desc')" }, includeTrashed: { type: "boolean", description: "Include trashed files (default false)" } } } } },
    { type: "function", function: { name: "get_drive_file", description: "Get metadata/details for a specific Drive file or folder.", parameters: { type: "object", properties: { fileId: { type: "string", description: "Drive file ID" } }, required: ["fileId"] } } },
    { type: "function", function: { name: "create_drive_folder", description: "Create a new folder in Google Drive.", parameters: { type: "object", properties: { name: { type: "string", description: "Folder name" }, parentId: { type: "string", description: "Optional parent folder ID" } }, required: ["name"] } } },
    { type: "function", function: { name: "create_drive_file", description: "Create a text file in Drive with optional parent folder.", parameters: { type: "object", properties: { name: { type: "string", description: "File name" }, content: { type: "string", description: "Text content" }, mimeType: { type: "string", description: "MIME type (default text/plain)" }, parentId: { type: "string", description: "Optional parent folder ID" } }, required: ["name"] } } },
    { type: "function", function: { name: "update_drive_file", description: "Update file content and/or rename a Drive file.", parameters: { type: "object", properties: { fileId: { type: "string", description: "Drive file ID" }, content: { type: "string", description: "New text content" }, name: { type: "string", description: "New file name" }, mimeType: { type: "string", description: "MIME type for content uploads (default text/plain)" } }, required: ["fileId"] } } },
    { type: "function", function: { name: "delete_drive_file", description: "Trash or permanently delete a Drive file.", parameters: { type: "object", properties: { fileId: { type: "string", description: "Drive file ID" }, permanent: { type: "boolean", description: "If true, permanently delete. Otherwise move to trash." } }, required: ["fileId"] } } },
    { type: "function", function: { name: "copy_drive_file", description: "Copy an existing Drive file.", parameters: { type: "object", properties: { fileId: { type: "string", description: "Source Drive file ID" }, name: { type: "string", description: "Optional new file name" }, parentId: { type: "string", description: "Optional parent folder for copy" } }, required: ["fileId"] } } },
    { type: "function", function: { name: "move_drive_file", description: "Move a Drive file to another folder.", parameters: { type: "object", properties: { fileId: { type: "string", description: "Drive file ID" }, newParentId: { type: "string", description: "Destination folder ID" } }, required: ["fileId", "newParentId"] } } },
    { type: "function", function: { name: "share_drive_file", description: "Share a Drive file with a user email.", parameters: { type: "object", properties: { fileId: { type: "string", description: "Drive file ID" }, emailAddress: { type: "string", description: "User email address to share with" }, role: { type: "string", description: "Permission role: reader, commenter, writer, organizer, fileOrganizer (default reader)" }, sendNotificationEmail: { type: "boolean", description: "Send share email notification (default true)" } }, required: ["fileId", "emailAddress"] } } },
    { type: "function", function: { name: "download_drive_file", description: "Download readable content for a Drive file (truncated for very large files).", parameters: { type: "object", properties: { fileId: { type: "string", description: "Drive file ID" }, maxBytes: { type: "integer", description: "Maximum bytes to return (default 200000)" } }, required: ["fileId"] } } }
];

const sheetsTools = [
    { type: "function", function: { name: "list_spreadsheets", description: "List spreadsheets you can access in Google Drive.", parameters: { type: "object", properties: { query: { type: "string", description: "Optional Drive query filter" }, maxResults: { type: "integer", description: "Max spreadsheets to return (default 25)" } } } } },
    { type: "function", function: { name: "create_spreadsheet", description: "Create a new Google Spreadsheet with optional sheet tab names.", parameters: { type: "object", properties: { title: { type: "string", description: "Spreadsheet title" }, sheets: { type: "array", items: { type: "string" }, description: "Optional sheet tab titles" } }, required: ["title"] } } },
    { type: "function", function: { name: "get_spreadsheet", description: "Get spreadsheet metadata and sheet tab info.", parameters: { type: "object", properties: { spreadsheetId: { type: "string", description: "Spreadsheet ID" }, includeGridData: { type: "boolean", description: "Include cell grid data (default false)" } }, required: ["spreadsheetId"] } } },
    { type: "function", function: { name: "list_sheet_tabs", description: "List all tab sheets in a spreadsheet.", parameters: { type: "object", properties: { spreadsheetId: { type: "string", description: "Spreadsheet ID" } }, required: ["spreadsheetId"] } } },
    { type: "function", function: { name: "add_sheet_tab", description: "Add a new tab sheet to a spreadsheet.", parameters: { type: "object", properties: { spreadsheetId: { type: "string", description: "Spreadsheet ID" }, title: { type: "string", description: "New tab title" }, rows: { type: "integer", description: "Initial row count (default 1000)" }, columns: { type: "integer", description: "Initial column count (default 26)" } }, required: ["spreadsheetId", "title"] } } },
    { type: "function", function: { name: "delete_sheet_tab", description: "Delete a tab sheet from a spreadsheet by sheetId.", parameters: { type: "object", properties: { spreadsheetId: { type: "string", description: "Spreadsheet ID" }, sheetId: { type: "integer", description: "Numeric sheet ID" } }, required: ["spreadsheetId", "sheetId"] } } },
    { type: "function", function: { name: "read_sheet_values", description: "Read values from a spreadsheet range (A1 notation).", parameters: { type: "object", properties: { spreadsheetId: { type: "string", description: "Spreadsheet ID" }, range: { type: "string", description: "A1 range (e.g. Sheet1!A1:D20)" }, valueRenderOption: { type: "string", description: "FORMATTED_VALUE, UNFORMATTED_VALUE, or FORMULA" }, dateTimeRenderOption: { type: "string", description: "SERIAL_NUMBER or FORMATTED_STRING" } }, required: ["spreadsheetId", "range"] } } },
    { type: "function", function: { name: "update_sheet_values", description: "Overwrite values in a spreadsheet range.", parameters: { type: "object", properties: { spreadsheetId: { type: "string", description: "Spreadsheet ID" }, range: { type: "string", description: "A1 range to write" }, values: { type: "array", description: "2D array of rows, e.g. [[\"Name\",\"Role\"],[\"Rishi\",\"Lead\"]]", items: { type: "array", items: { type: "string" } } }, valueInputOption: { type: "string", description: "RAW or USER_ENTERED (default USER_ENTERED)" }, majorDimension: { type: "string", description: "ROWS or COLUMNS (default ROWS)" } }, required: ["spreadsheetId", "range", "values"] } } },
    { type: "function", function: { name: "update_timesheet_hours", description: "Find a timesheet row by date and reliably update one or more fields (billing hours, task details, non-billing hours, project/module/month).", parameters: { type: "object", properties: { spreadsheetId: { type: "string", description: "Spreadsheet ID" }, sheetName: { type: "string", description: "Tab name (default Tracker)" }, date: { type: "string", description: "Date to match, e.g. 6-Feb-2026 or 2026-02-06" }, billingHours: { type: "number", description: "Billing hours value to set" }, taskDetails: { type: "string", description: "Task details/description to set" }, nonBillingHours: { type: "number", description: "Non-billing hours value to set" }, projectName: { type: "string", description: "Project name to set" }, moduleName: { type: "string", description: "Module name to set" }, month: { type: "string", description: "Month label to set (e.g. February 2026)" }, dateColumn: { type: "string", description: "Date column letter (default B)" }, taskDetailsColumn: { type: "string", description: "Task Details column letter (default C)" }, billingHoursColumn: { type: "string", description: "Billing hours column letter (default D)" }, nonBillingHoursColumn: { type: "string", description: "Non-billing hours column letter (default E)" }, projectNameColumn: { type: "string", description: "Project name column letter (default F)" }, moduleNameColumn: { type: "string", description: "Module name column letter (default G)" }, monthColumn: { type: "string", description: "Month column letter (default A)" }, searchRange: { type: "string", description: "A1 range to search (default A1:Z3000)" }, preferEmptyBilling: { type: "boolean", description: "Prefer row where billing cell is empty when duplicates exist and billingHours is being updated (default true)" } }, required: ["spreadsheetId", "date"] } } },
    { type: "function", function: { name: "append_sheet_values", description: "Append rows to a spreadsheet range.", parameters: { type: "object", properties: { spreadsheetId: { type: "string", description: "Spreadsheet ID" }, range: { type: "string", description: "A1 target range (e.g. Sheet1!A:D)" }, values: { type: "array", description: "2D array of rows to append", items: { type: "array", items: { type: "string" } } }, valueInputOption: { type: "string", description: "RAW or USER_ENTERED (default USER_ENTERED)" }, insertDataOption: { type: "string", description: "INSERT_ROWS or OVERWRITE (default INSERT_ROWS)" } }, required: ["spreadsheetId", "range", "values"] } } },
    { type: "function", function: { name: "clear_sheet_values", description: "Clear values in a spreadsheet range.", parameters: { type: "object", properties: { spreadsheetId: { type: "string", description: "Spreadsheet ID" }, range: { type: "string", description: "A1 range to clear" } }, required: ["spreadsheetId", "range"] } } }
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

const outlookTools = [
    { type: "function", function: { name: "outlook_send_email", description: "Send an email via Outlook. Supports CC/BCC. Body can be plain text or HTML.", parameters: { type: "object", properties: { to: { type: "array", items: { type: "string" }, description: "Recipient email addresses" }, subject: { type: "string", description: "Email subject line" }, body: { type: "string", description: "Email body (plain text or HTML)" }, cc: { type: "array", items: { type: "string" }, description: "CC addresses" }, bcc: { type: "array", items: { type: "string" }, description: "BCC addresses" } }, required: ["to", "subject", "body"] } } },
    { type: "function", function: { name: "outlook_list_emails", description: "List recent emails from an Outlook folder (default: inbox).", parameters: { type: "object", properties: { maxResults: { type: "integer", description: "Number of emails to return (default 20)" }, folder: { type: "string", description: "Folder name: inbox, sentitems, drafts, deleteditems, junkemail (default: inbox)" } } } } },
    { type: "function", function: { name: "outlook_read_email", description: "Read full content of an Outlook email by message ID.", parameters: { type: "object", properties: { messageId: { type: "string", description: "Outlook message ID" } }, required: ["messageId"] } } },
    { type: "function", function: { name: "outlook_search_emails", description: "Search Outlook emails using Microsoft Search syntax (supports KQL: from:, subject:, hasAttachment:true, etc.).", parameters: { type: "object", properties: { query: { type: "string", description: "Search query (e.g., 'from:john subject:meeting')" }, maxResults: { type: "integer", description: "Number of results (default 20)" } }, required: ["query"] } } },
    { type: "function", function: { name: "outlook_reply_to_email", description: "Reply to an Outlook email.", parameters: { type: "object", properties: { messageId: { type: "string", description: "Message ID to reply to" }, body: { type: "string", description: "Reply body (HTML or text)" } }, required: ["messageId", "body"] } } },
    { type: "function", function: { name: "outlook_forward_email", description: "Forward an Outlook email to new recipients.", parameters: { type: "object", properties: { messageId: { type: "string", description: "Message ID to forward" }, to: { type: "array", items: { type: "string" }, description: "Forward recipient addresses" }, comment: { type: "string", description: "Comment to include" } }, required: ["messageId", "to"] } } },
    { type: "function", function: { name: "outlook_delete_email", description: "Delete an Outlook email permanently.", parameters: { type: "object", properties: { messageId: { type: "string", description: "Message ID to delete" } }, required: ["messageId"] } } },
    { type: "function", function: { name: "outlook_move_email", description: "Move an Outlook email to a different folder.", parameters: { type: "object", properties: { messageId: { type: "string", description: "Message ID to move" }, destinationFolderId: { type: "string", description: "Destination folder ID (use outlook_list_folders to find IDs)" } }, required: ["messageId", "destinationFolderId"] } } },
    { type: "function", function: { name: "outlook_mark_as_read", description: "Mark an Outlook email as read.", parameters: { type: "object", properties: { messageId: { type: "string", description: "Message ID" } }, required: ["messageId"] } } },
    { type: "function", function: { name: "outlook_mark_as_unread", description: "Mark an Outlook email as unread.", parameters: { type: "object", properties: { messageId: { type: "string", description: "Message ID" } }, required: ["messageId"] } } },
    { type: "function", function: { name: "outlook_list_folders", description: "List all Outlook mail folders with item counts.", parameters: { type: "object", properties: {} } } },
    { type: "function", function: { name: "outlook_create_folder", description: "Create a new Outlook mail folder.", parameters: { type: "object", properties: { name: { type: "string", description: "Folder display name" }, parentFolderId: { type: "string", description: "Parent folder ID (omit for top-level)" } }, required: ["name"] } } },
    { type: "function", function: { name: "outlook_get_attachments", description: "List attachments on an Outlook email.", parameters: { type: "object", properties: { messageId: { type: "string", description: "Message ID" } }, required: ["messageId"] } } },
    { type: "function", function: { name: "outlook_create_draft", description: "Create a draft email in Outlook.", parameters: { type: "object", properties: { to: { type: "array", items: { type: "string" }, description: "Recipient addresses" }, subject: { type: "string", description: "Subject line" }, body: { type: "string", description: "Email body (HTML or text)" }, cc: { type: "array", items: { type: "string" }, description: "CC addresses" }, bcc: { type: "array", items: { type: "string" }, description: "BCC addresses" } } } } },
    { type: "function", function: { name: "outlook_send_draft", description: "Send an existing Outlook draft.", parameters: { type: "object", properties: { messageId: { type: "string", description: "Draft message ID" } }, required: ["messageId"] } } },
    { type: "function", function: { name: "outlook_list_drafts", description: "List Outlook draft emails.", parameters: { type: "object", properties: { maxResults: { type: "integer", description: "Number of drafts to return (default 20)" } } } } },
    { type: "function", function: { name: "outlook_flag_email", description: "Set or clear a flag on an Outlook email.", parameters: { type: "object", properties: { messageId: { type: "string", description: "Message ID" }, flagStatus: { type: "string", description: "Flag status: flagged, complete, notFlagged (default: flagged)" } }, required: ["messageId"] } } },
    { type: "function", function: { name: "outlook_get_user_profile", description: "Get the connected Outlook/Microsoft user's profile information.", parameters: { type: "object", properties: {} } } }
];

const docsTools = [
    { type: "function", function: { name: "list_documents", description: "List Google Docs documents. Uses Drive API to find documents by name.", parameters: { type: "object", properties: { query: { type: "string", description: "Optional name filter (e.g. 'Meeting Notes')" }, pageSize: { type: "integer", description: "Max documents to return (default 25)" } } } } },
    { type: "function", function: { name: "get_document", description: "Get full structure/metadata of a Google Doc by document ID.", parameters: { type: "object", properties: { documentId: { type: "string", description: "Google Doc document ID" } }, required: ["documentId"] } } },
    { type: "function", function: { name: "create_document", description: "Create a new Google Doc with a title and optional initial content.", parameters: { type: "object", properties: { title: { type: "string", description: "Document title" }, content: { type: "string", description: "Optional initial text content" } }, required: ["title"] } } },
    { type: "function", function: { name: "insert_text", description: "Insert text at a specific index in a Google Doc.", parameters: { type: "object", properties: { documentId: { type: "string", description: "Google Doc document ID" }, text: { type: "string", description: "Text to insert" }, index: { type: "integer", description: "Character index to insert at (default 1 = start)" } }, required: ["documentId", "text"] } } },
    { type: "function", function: { name: "replace_text", description: "Find and replace all occurrences of text in a Google Doc.", parameters: { type: "object", properties: { documentId: { type: "string", description: "Google Doc document ID" }, findText: { type: "string", description: "Text to find" }, replaceWith: { type: "string", description: "Replacement text" }, matchCase: { type: "boolean", description: "Case-sensitive match (default false)" } }, required: ["documentId", "findText", "replaceWith"] } } },
    { type: "function", function: { name: "delete_content", description: "Delete content in a Google Doc between start and end indices.", parameters: { type: "object", properties: { documentId: { type: "string", description: "Google Doc document ID" }, startIndex: { type: "integer", description: "Start character index" }, endIndex: { type: "integer", description: "End character index" } }, required: ["documentId", "startIndex", "endIndex"] } } },
    { type: "function", function: { name: "append_text", description: "Append text to the end of a Google Doc.", parameters: { type: "object", properties: { documentId: { type: "string", description: "Google Doc document ID" }, text: { type: "string", description: "Text to append" } }, required: ["documentId", "text"] } } },
    { type: "function", function: { name: "get_document_text", description: "Extract plain text content from a Google Doc.", parameters: { type: "object", properties: { documentId: { type: "string", description: "Google Doc document ID" } }, required: ["documentId"] } } }
];

const teamsTools = [
    { type: "function", function: { name: "teams_list_teams", description: "List Microsoft Teams the user has joined.", parameters: { type: "object", properties: {} } } },
    { type: "function", function: { name: "teams_get_team", description: "Get details about a specific Microsoft Team.", parameters: { type: "object", properties: { teamId: { type: "string", description: "Team ID" } }, required: ["teamId"] } } },
    { type: "function", function: { name: "teams_list_channels", description: "List channels in a Microsoft Team.", parameters: { type: "object", properties: { teamId: { type: "string", description: "Team ID" } }, required: ["teamId"] } } },
    { type: "function", function: { name: "teams_send_channel_message", description: "Send a message to a channel in a Microsoft Team.", parameters: { type: "object", properties: { teamId: { type: "string", description: "Team ID" }, channelId: { type: "string", description: "Channel ID" }, message: { type: "string", description: "Message content (plain text or HTML)" }, contentType: { type: "string", description: "Content type: text or html (default text)" } }, required: ["teamId", "channelId", "message"] } } },
    { type: "function", function: { name: "teams_list_channel_messages", description: "List recent messages from a channel in a Microsoft Team.", parameters: { type: "object", properties: { teamId: { type: "string", description: "Team ID" }, channelId: { type: "string", description: "Channel ID" }, top: { type: "integer", description: "Number of messages to return (default 20)" } }, required: ["teamId", "channelId"] } } },
    { type: "function", function: { name: "teams_list_chats", description: "List the user's Microsoft Teams chats (1:1 and group).", parameters: { type: "object", properties: { top: { type: "integer", description: "Number of chats to return (default 20)" } } } } },
    { type: "function", function: { name: "teams_send_chat_message", description: "Send a message in a Microsoft Teams chat.", parameters: { type: "object", properties: { chatId: { type: "string", description: "Chat ID" }, message: { type: "string", description: "Message content (plain text or HTML)" }, contentType: { type: "string", description: "Content type: text or html (default text)" } }, required: ["chatId", "message"] } } },
    { type: "function", function: { name: "teams_list_chat_messages", description: "List recent messages from a Microsoft Teams chat.", parameters: { type: "object", properties: { chatId: { type: "string", description: "Chat ID" }, top: { type: "integer", description: "Number of messages to return (default 20)" } }, required: ["chatId"] } } },
    { type: "function", function: { name: "teams_create_chat", description: "Create a new Microsoft Teams 1:1 or group chat.", parameters: { type: "object", properties: { chatType: { type: "string", description: "Chat type: oneOnOne or group" }, members: { type: "array", items: { type: "string" }, description: "Array of member email addresses" }, topic: { type: "string", description: "Chat topic (for group chats)" } }, required: ["chatType", "members"] } } },
    { type: "function", function: { name: "teams_get_chat_members", description: "List members of a Microsoft Teams chat.", parameters: { type: "object", properties: { chatId: { type: "string", description: "Chat ID" } }, required: ["chatId"] } } }
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
const teamsToolNames = new Set(teamsTools.map(t => t.function.name));

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
        download_drive_file: downloadDriveFile
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

// Master dispatcher
async function executeTool(toolName, args) {
    console.log(`[Tool] ${toolName}`, JSON.stringify(args).slice(0, 200));
    if (gmailToolNames.has(toolName)) return await executeGmailTool(toolName, args);
    if (calendarToolNames.has(toolName)) return await executeCalendarTool(toolName, args);
    if (gchatToolNames.has(toolName)) return await executeGchatTool(toolName, args);
    if (driveToolNames.has(toolName)) return await executeDriveTool(toolName, args);
    if (sheetsToolNames.has(toolName)) return await executeSheetsTool(toolName, args);
    if (toolName.startsWith(SHEETS_MCP_TOOL_PREFIX)) return await executeSheetsMcpTool(toolName, args);
    if (docsToolNames.has(toolName)) return await executeDocsTool(toolName, args);
    if (githubToolNames.has(toolName)) return await executeGitHubTool(toolName, args);
    if (outlookToolNames.has(toolName)) return await executeOutlookTool(toolName, args);
    if (teamsToolNames.has(toolName)) return await executeTeamsTool(toolName, args);
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
    const docs = { service: 'docs', connected: docsConnected, tools: docsTools.map(t => ({ function: t.function })) };
    const teamsConnected = !!outlookAccessToken && hasTeamsScopes();
    const teams = { service: 'teams', connected: teamsConnected, tools: teamsTools.map(t => ({ function: t.function })) };
    const totalTools = gmailTools.length + calendarTools.length + gchatTools.length + driveTools.length + sheetsTools.length + sheetsMcpTools.length + githubTools.length + outlookTools.length + docsTools.length + teamsTools.length;
    res.json({
        services: [gmail, calendar, gchat, drive, sheets, github, outlook, docs, teams],
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
        scheduledTasks.push(task);
        saveScheduledTasksToDisk();
        res.status(201).json({ success: true, task });
    } catch (error) {
        res.status(400).json({ error: error.message });
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
        res.status(400).json({ error: error.message });
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
    const task = scheduledTasks.find(item => item.id === req.params.id);
    if (!task) return res.status(404).json({ error: 'Task not found' });
    const result = await runScheduledTask(task.id, 'manual');
    if (!result.success && !result.skipped) {
        return res.status(500).json(result);
    }
    return res.json(result);
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

// Google Chat authentication status (same Google OAuth, separate scopes)
app.get('/api/gchat/status', (req, res) => {
    const hasCredentials = (process.env.GMAIL_CLIENT_ID || process.env.GOOGLE_CLIENT_ID) &&
        (process.env.GMAIL_CLIENT_SECRET || process.env.GOOGLE_CLIENT_SECRET);
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
    const hasCredentials = (process.env.GMAIL_CLIENT_ID || process.env.GOOGLE_CLIENT_ID) &&
        (process.env.GMAIL_CLIENT_SECRET || process.env.GOOGLE_CLIENT_SECRET);
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

// Google Sheets authentication status (same Google OAuth, separate scope)
app.get('/api/sheets/status', (req, res) => {
    const hasCredentials = (process.env.GMAIL_CLIENT_ID || process.env.GOOGLE_CLIENT_ID) &&
        (process.env.GMAIL_CLIENT_SECRET || process.env.GOOGLE_CLIENT_SECRET);
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
        command: SHEETS_MCP_COMMAND,
        args: SHEETS_MCP_ARGS,
        credentialsDirectory: SHEETS_MCP_CREDS_DIR,
        error: sheetsMcpError
    });
});

app.post('/api/sheets-mcp/reconnect', async (req, res) => {
    await closeSheetsMcpClient();
    await initSheetsMcpClient();
    res.json({
        success: !!sheetsMcpClient,
        connected: !!sheetsMcpClient,
        toolCount: sheetsMcpTools.length,
        error: sheetsMcpError
    });
});

// Google Docs status (same Google OAuth, separate scope)
app.get('/api/docs/status', (req, res) => {
    const hasCredentials = (process.env.GMAIL_CLIENT_ID || process.env.GOOGLE_CLIENT_ID) &&
        (process.env.GMAIL_CLIENT_SECRET || process.env.GOOGLE_CLIENT_SECRET);
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
    const authUrl = oauth2Client.generateAuthUrl({ access_type: 'offline', scope: SCOPES, prompt: GOOGLE_OAUTH_PROMPT });
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
    const authUrl = oauth2Client.generateAuthUrl({ access_type: 'offline', scope: SCOPES, prompt: GOOGLE_OAUTH_PROMPT });
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
    const authUrl = oauth2Client.generateAuthUrl({ access_type: 'offline', scope: SCOPES, prompt: GOOGLE_OAUTH_PROMPT });
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
    const authUrl = oauth2Client.generateAuthUrl({ access_type: 'offline', scope: SCOPES, prompt: GOOGLE_OAUTH_PROMPT });
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
        res.status(401).json({ error: `Invalid token: ${error.message}` });
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
    const authUrl = oauth2Client.generateAuthUrl({ access_type: 'offline', scope: SCOPES, prompt: GOOGLE_OAUTH_PROMPT });
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
            return res.status(401).send(`GitHub token exchange failed: ${details}`);
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
        res.send(`<html><body style="background:#0f0f1a;color:#fff;display:flex;align-items:center;justify-content:center;height:100vh;font-family:sans-serif"><div style="text-align:center"><h1>Google Connected!</h1><p>${calendarMessage} ${gchatMessage} ${driveMessage} ${sheetsMessage} You can close this window.</p></div></body></html>`);
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
            return res.status(401).send(`Outlook token exchange failed: ${details}`);
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
    if (docsConnected) availableTools.push(...docsTools);
    if (octokitClient) availableTools.push(...githubTools);
    if (outlookAccessToken) availableTools.push(...outlookTools);
    const teamsConnected = !!outlookAccessToken && hasTeamsScopes();
    if (teamsConnected) availableTools.push(...teamsTools);

    const connectedServices = [];
    if (gmailClient) connectedServices.push('Gmail (25 tools)');
    if (calendarConnected) connectedServices.push(`Google Calendar (${calendarTools.length} tools)`);
    if (gchatConnected) connectedServices.push(`Google Chat (${gchatTools.length} tools)`);
    if (driveConnected) connectedServices.push(`Google Drive (${driveTools.length} tools)`);
    if (sheetsConnected) connectedServices.push(`Google Sheets (${sheetsTools.length} tools)`);
    if (sheetsMcpConnected) connectedServices.push(`Google Sheets MCP (${sheetsMcpTools.length} tools)`);
    if (docsConnected) connectedServices.push(`Google Docs (${docsTools.length} tools)`);
    if (octokitClient) connectedServices.push('GitHub (20 tools)');
    if (outlookAccessToken) connectedServices.push(`Outlook (${outlookTools.length} tools)`);
    if (teamsConnected) connectedServices.push(`Microsoft Teams (${teamsTools.length} tools)`);

    const statusText = connectedServices.length > 0 ? connectedServices.join(', ') : 'No services connected';
    return { availableTools, statusText };
}

function buildAgentSystemPrompt({ statusText, toolCount, dateContext }) {
    const basePrompt = `You are a powerful AI assistant with tools across Gmail, Google Calendar, Google Chat, Google Drive, Google Sheets, Google Docs, GitHub, Outlook, and Microsoft Teams. You can perform complex, multi-step operations across all connected services.

Connected Services: ${statusText}
Total Tools Available: ${toolCount}

## EXECUTION MODE:
- A single user command can contain multiple tasks. Treat it as one end-to-end workflow, not separate mini chats.
- Create an internal plan, execute tools, evaluate results, and continue until the user goal is complete or truly blocked.
- Prefer action over clarification. Ask follow-up questions only when a required value is missing or an operation is destructive.

## CORE RULES - Follow these STRICTLY:

1. **DISCOVERY FIRST, NEVER GUESS**: When the user refers to emails/docs/files/issues by description, use search/list/discovery tools first. Never invent IDs, email addresses, or repository names.

2. **ONE COMMAND -> MULTI-TOOL EXECUTION**: If the user asks for a compound task, execute all required steps in the same request flow.
   - "Read and reply to John's latest email" -> search_emails -> read_email -> reply_to_email
   - "Create an issue for the bug and email the team" -> create_issue -> send_email
   - "Find calendar events for today and email attendees" -> list_events -> send_email
   - "List my spreadsheets and summarize tabs" -> list_spreadsheets -> list_sheet_tabs
   - "Check my PRs and schedule review slots" -> list_pull_requests -> create_event

3. **PARALLELIZE WHEN INDEPENDENT**: When subtasks do not depend on each other, call multiple tools in the same turn.
   - Example: gather Gmail unread count, Outlook unread count, and today's calendar events at the same time.

4. **SEQUENCE WHEN DEPENDENT**: If a later step needs output from an earlier step, chain tools in order.
   - Example: search -> read -> extract recipient -> draft/send.

5. **NEVER STOP MID-TASK**: Continue tool execution until completion. Do not ask for confirmation between non-destructive intermediate steps.

6. **RECOVER FROM FAILURES**: If a tool fails due to missing ID, scope, or lookup ambiguity, run the appropriate discovery/diagnostic tool and retry with corrected arguments. Explain blockers only after retry paths are exhausted.

7. **SAFETY FOR DESTRUCTIVE ACTIONS**: For delete/trash/clear/bulk-destructive operations, ask for confirmation unless the user explicitly requested that exact destructive action.

8. **USE BATCH OPERATIONS**: For multi-email modifications, prefer batch_modify_emails over repeated single-item calls.

9. **CROSS-SERVICE ORCHESTRATION**: Combine Gmail, Calendar, Chat, Drive, Sheets, Docs, GitHub, Outlook, and Teams tools when useful to finish the full request.

10. **BE PROACTIVE BUT ACCURATE**: Share helpful observations discovered during execution, but never claim success unless tool output confirms it.

## TOOL USAGE TIPS:
- Gmail: search_emails supports full Gmail query syntax (from:, to:, subject:, is:unread, has:attachment, etc.)
- Calendar: Use list_events with timeMin/timeMax for date ranges. Use create_meet_event or create_event with createMeetLink=true for Google Meet.
- Calendar attendees: If the user gives a person name (not exact email), resolve it from Gmail history first and do not use placeholder domains like example.com.
- Calendar availability: Use check_person_availability for one person and find_common_free_slots for multiple people. Availability only works for calendars the user can access.
- Calendar IDs: get_event requires a Calendar eventId, not a Gmail messageId from search_emails/read_email.
- Google Chat: Use list_chat_spaces to discover spaces, then send_chat_message to post updates.
- Drive: Use list_drive_files for discovery before updates/deletes. Use share_drive_file to grant access.
- Sheets: Use list_spreadsheets then list_sheet_tabs/read_sheet_values before edits. Use update_sheet_values/append_sheet_values for writes.
- Sheets timesheets: For any date-based timesheet update (hours and/or task details), ALWAYS use update_timesheet_hours (not update_sheet_values/append_sheet_values), and pass the user's exact date phrase (e.g., "Feb 6th 2026").
- Sheets timesheets row safety: Resolve the row by date and update that row only. Never hardcode row numbers/cell addresses from prior runs.
- Sheets MCP: MCP tools are prefixed with ${SHEETS_MCP_TOOL_PREFIX} and provide fallback read/update operations when needed.
- Sheets write safety: Never claim a sheet update succeeded unless tool output indicates verification/read-back succeeded.
- GitHub: Use owner/repo format. search_repos for discovery. list_issues/list_pull_requests for project management.
- Outlook: All Outlook tools are prefixed with outlook_. outlook_search_emails supports Microsoft KQL (from:, subject:, hasAttachment:true). Use outlook_list_folders to get folder IDs before moving emails. Cross-service: combine Gmail + Outlook for multi-mailbox workflows.
- Google Docs: Use list_documents to find Docs by name. Use get_document_text to read content. Use create_document to create new docs. Use insert_text/append_text to add content. Use replace_text for find-and-replace.
- Microsoft Teams: All Teams tools are prefixed with teams_. Use teams_list_teams to discover teams, then teams_list_channels for channels. Use teams_send_channel_message to post. Use teams_list_chats for 1:1/group chats, teams_send_chat_message to reply.

## FINAL RESPONSE QUALITY:
- Provide a concise outcome summary of what was completed.
- Include key outputs/links/IDs the user needs next.
- If anything failed, state exactly what failed and what retry/permission is needed.
- Do not include internal chain-of-thought. Keep reasoning brief and action-focused.`;

    const dateContextPrompt = `DATE CONTEXT FOR THIS REQUEST
- Current timestamp (UTC): ${dateContext.nowIso}
- Local timezone: ${dateContext.timeZone}
- Today: ${dateContext.today} (${dateContext.weekday})
- Tomorrow: ${dateContext.tomorrow}
- Yesterday: ${dateContext.yesterday}

Rules:
- Resolve relative date words (today, tomorrow, yesterday, this week, next week) using this context.
- For date-specific calendar lookups, always send explicit ISO timeMin and timeMax to list_events, get_free_busy, check_person_availability, and find_common_free_slots.`;

    return `${basePrompt}\n\n${dateContextPrompt}`;
}

async function runAgentConversation({ message, history = [] }) {
    if (!process.env.OPENAI_API_KEY) {
        throw new Error('OpenAI API key missing');
    }

    const { availableTools, statusText } = getConnectedToolContext();
    const dateContext = getCurrentDateContext();
    const systemPrompt = buildAgentSystemPrompt({
        statusText,
        toolCount: availableTools.length,
        dateContext
    });
    const enrichedUserMessage = `${message}

[Runtime Date Context]
- Current timestamp (UTC): ${dateContext.nowIso}
- Local timezone: ${dateContext.timeZone}
- Today: ${dateContext.today}
- Tomorrow: ${dateContext.tomorrow}
- Yesterday: ${dateContext.yesterday}`;

    const messages = [
        { role: 'system', content: systemPrompt },
        ...history,
        { role: 'user', content: enrichedUserMessage }
    ];

    let response = await openai.chat.completions.create({
        model: 'gpt-4o',
        messages,
        tools: availableTools.length > 0 ? availableTools : undefined,
        tool_choice: availableTools.length > 0 ? 'auto' : undefined
    });

    let assistantMessage = response.choices[0].message;
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
                args = JSON.parse(toolCall.function.arguments);
            } catch (error) {
                return { toolCall, error: `Invalid arguments: ${error.message}` };
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
            allToolResults.push(error
                ? { tool: toolCall.function.name, error }
                : { tool: toolCall.function.name, result });

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

    return {
        response: finalResponse,
        toolResults: allToolResults,
        steps: allSteps,
        turnsUsed: turnCount
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
        const dueTasks = scheduledTasks.filter(task =>
            task.enabled &&
            task.time === now.hhmm &&
            task.lastRunDate !== now.date
        );
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
        res.status(500).json({ error: error.message });
    }
});

app.post('/api/chat', async (req, res) => {
    const { message, history = [] } = req.body;
    try {
        const result = await runAgentConversation({ message, history });
        res.json(result);
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
initSheetsMcpClient().catch(error => {
    console.error('Sheets MCP bootstrap failed:', error?.message || error);
});
initOutlookClient().catch(error => {
    console.error('Outlook init failed:', error?.message || error);
});
loadScheduledTasksFromDisk();
startScheduledTaskRunner();

const totalTools = gmailTools.length + calendarTools.length + gchatTools.length + driveTools.length + sheetsTools.length + sheetsMcpTools.length + githubTools.length + outlookTools.length + docsTools.length + teamsTools.length;
app.listen(PORT, () => {
    console.log(`\nAI Agent Server running at http://localhost:${PORT}`);
    console.log(`Total tools available: ${totalTools} (Gmail: ${gmailTools.length}, Calendar: ${calendarTools.length}, Chat: ${gchatTools.length}, Drive: ${driveTools.length}, Sheets: ${sheetsTools.length}, Sheets MCP: ${sheetsMcpTools.length}, Docs: ${docsTools.length}, GitHub: ${githubTools.length}, Outlook: ${outlookTools.length}, Teams: ${teamsTools.length})`);
    console.log(`Gmail: ${gmailClient ? 'Connected' : 'Not connected'}`);
    console.log(`Calendar: ${calendarClient && hasCalendarScope() ? 'Connected' : 'Not connected'}`);
    console.log(`Google Chat: ${gchatClient && hasGchatScopes() ? 'Connected' : 'Not connected'}`);
    console.log(`Google Drive: ${driveClient && hasDriveScope() ? 'Connected' : 'Not connected'}`);
    console.log(`Google Sheets: ${sheetsClient && hasSheetsScope() ? 'Connected' : 'Not connected'}`);
    console.log(`Google Sheets MCP: ${sheetsMcpClient ? `Connected (${sheetsMcpTools.length} tools)` : `Not connected${sheetsMcpError ? ` (${sheetsMcpError})` : ''}`}`);
    console.log(`Google Docs: ${docsClient && hasDocsScope() ? 'Connected' : 'Not connected'}`);
    console.log(`GitHub: ${octokitClient ? 'Connected' : 'Not connected'}`);
    console.log(`Outlook: ${outlookAccessToken ? `Connected (${outlookUserEmail || 'unknown'})` : 'Not connected'}`);
    console.log(`Microsoft Teams: ${outlookAccessToken && hasTeamsScopes() ? 'Connected' : 'Not connected'}`);
    console.log(`Timer Tasks: ${scheduledTasks.length} configured (${scheduledTasks.filter(task => task.enabled).length} enabled)`);
});

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




