// DOM Elements
const chatMessages = document.getElementById('chatMessages');
const messageInput = document.getElementById('messageInput');
const sendBtn = document.getElementById('sendBtn');
const gmailNavItem = document.getElementById('gmailNavItem');
const calendarNavItem = document.getElementById('calendarNavItem');
const gchatNavItem = document.getElementById('gchatNavItem');
const githubNavItem = document.getElementById('githubNavItem');
const gmailPanel = document.getElementById('gmailPanel');
const calendarPanel = document.getElementById('calendarPanel');
const gchatPanel = document.getElementById('gchatPanel');
const githubPanel = document.getElementById('githubPanel');
const authenticateBtn = document.getElementById('authenticateBtn');
const calendarAuthBtn = document.getElementById('calendarAuthBtn');
const gchatAuthBtn = document.getElementById('gchatAuthBtn');
const githubAuthBtn = document.getElementById('githubAuthBtn');
const githubDisconnectBtn = document.getElementById('githubDisconnectBtn');
const gmailReauthBtn = document.getElementById('gmailReauthBtn');
const calendarReauthBtn = document.getElementById('calendarReauthBtn');
const gchatReauthBtn = document.getElementById('gchatReauthBtn');
const githubReauthBtn = document.getElementById('githubReauthBtn');
const githubAuthNote = document.getElementById('githubAuthNote');
const authSection = document.getElementById('authSection');
const connectedSection = document.getElementById('connectedSection');
const setupSection = document.getElementById('setupSection');
const calendarAuthSection = document.getElementById('calendarAuthSection');
const calendarConnectedSection = document.getElementById('calendarConnectedSection');
const gchatAuthSection = document.getElementById('gchatAuthSection');
const gchatConnectedSection = document.getElementById('gchatConnectedSection');
const githubAuthSection = document.getElementById('githubAuthSection');
const githubConnectedSection = document.getElementById('githubConnectedSection');
const gmailStatus = document.getElementById('gmailStatus');
const calendarStatus = document.getElementById('calendarStatus');
const gchatStatus = document.getElementById('gchatStatus');
const githubStatus = document.getElementById('githubStatus');
const gmailBadge = document.getElementById('gmailBadge');
const calendarBadge = document.getElementById('calendarBadge');
const gchatBadge = document.getElementById('gchatBadge');
const githubBadge = document.getElementById('githubBadge');
const turnsBadge = document.getElementById('turnsBadge');
const turnsCount = document.getElementById('turnsCount');
const capabilitiesNavItem = document.getElementById('capabilitiesNavItem');
const capabilitiesModal = document.getElementById('capabilitiesModal');
const closeCapabilitiesBtn = document.getElementById('closeCapabilitiesBtn');
const toolCountBadge = document.getElementById('toolCountBadge');
const toolStatusText = document.getElementById('toolStatusText');
const modalTitle = document.getElementById('modalTitle');

// State
let chatHistory = [];
let isGmailConnected = false;
let isCalendarConnected = false;
let isGchatConnected = false;
let isGithubConnected = false;
let activeFilter = 'all';

// Tool icon mapping
const TOOL_ICONS = {
    // Gmail (25)
    send_email: '&#9993;', search_emails: '&#128269;', read_email: '&#128214;',
    list_emails: '&#128203;', trash_email: '&#128465;', modify_labels: '&#127991;',
    create_draft: '&#128221;', reply_to_email: '&#8617;', forward_email: '&#10145;',
    list_labels: '&#127991;', create_label: '&#10133;', delete_label: '&#10134;',
    mark_as_read: '&#9745;', mark_as_unread: '&#9746;', star_email: '&#11088;',
    unstar_email: '&#9734;', archive_email: '&#128230;', untrash_email: '&#9854;',
    get_thread: '&#128172;', list_drafts: '&#128196;', delete_draft: '&#128465;',
    send_draft: '&#128228;', get_attachment_info: '&#128206;', get_profile: '&#128100;',
    batch_modify_emails: '&#9881;',
    // Calendar
    list_events: '&#128197;', get_event: '&#128196;', create_event: '&#10133;',
    create_meet_event: '&#128249;', add_meet_link_to_event: '&#128279;',
    update_event: '&#9998;', delete_event: '&#128465;', list_calendars: '&#128197;',
    create_calendar: '&#10133;', quick_add_event: '&#9889;', get_free_busy: '&#128338;',
    check_person_availability: '&#128100;', find_common_free_slots: '&#129309;',
    list_recurring_instances: '&#128257;', move_event: '&#10145;',
    update_event_attendees: '&#128101;', get_calendar_colors: '&#127912;',
    clear_calendar: '&#128465;', watch_events: '&#128276;',
    // Google Chat (3)
    list_chat_spaces: '&#128172;', send_chat_message: '&#9993;', list_chat_messages: '&#128221;',
    // GitHub (20)
    list_repos: '&#128193;', get_repo: '&#128196;', create_repo: '&#10133;',
    list_issues: '&#128196;', create_issue: '&#10133;', update_issue: '&#9998;',
    list_pull_requests: '&#128259;', get_pull_request: '&#128196;',
    create_pull_request: '&#10133;', merge_pull_request: '&#128279;',
    list_branches: '&#128204;', create_branch: '&#128204;',
    get_file_content: '&#128196;', create_or_update_file: '&#128196;',
    search_repos: '&#128269;', search_code: '&#128269;',
    list_commits: '&#128221;', get_user_profile: '&#128100;',
    list_notifications: '&#128276;', list_gists: '&#128221;'
};

// Tool category mappings per service
const GMAIL_CATEGORIES = {
    'Core': ['send_email', 'search_emails', 'read_email', 'list_emails'],
    'Actions': ['reply_to_email', 'forward_email', 'trash_email', 'untrash_email', 'archive_email'],
    'Status': ['mark_as_read', 'mark_as_unread', 'star_email', 'unstar_email'],
    'Labels': ['list_labels', 'create_label', 'delete_label', 'modify_labels'],
    'Drafts': ['create_draft', 'list_drafts', 'delete_draft', 'send_draft'],
    'Advanced': ['get_thread', 'get_attachment_info', 'get_profile', 'batch_modify_emails']
};

const CALENDAR_CATEGORIES = {
    'Events': ['list_events', 'get_event', 'create_event', 'update_event', 'delete_event', 'quick_add_event'],
    'Meet': ['create_meet_event', 'add_meet_link_to_event'],
    'Calendars': ['list_calendars', 'create_calendar', 'clear_calendar'],
    'Scheduling': ['get_free_busy', 'check_person_availability', 'find_common_free_slots', 'list_recurring_instances', 'move_event'],
    'Management': ['update_event_attendees', 'get_calendar_colors', 'watch_events']
};

const GCHAT_CATEGORIES = {
    'Spaces': ['list_chat_spaces'],
    'Messages': ['send_chat_message', 'list_chat_messages']
};

const GITHUB_CATEGORIES = {
    'Repositories': ['list_repos', 'get_repo', 'create_repo', 'search_repos'],
    'Issues': ['list_issues', 'create_issue', 'update_issue'],
    'Pull Requests': ['list_pull_requests', 'get_pull_request', 'create_pull_request', 'merge_pull_request'],
    'Code & Branches': ['list_branches', 'create_branch', 'get_file_content', 'create_or_update_file', 'search_code'],
    'Activity': ['list_commits', 'get_user_profile', 'list_notifications', 'list_gists']
};

// Initialize
document.addEventListener('DOMContentLoaded', () => {
    checkAllStatuses();
    setupEventListeners();
    autoResizeTextarea();
});

// Event Listeners
function setupEventListeners() {
    sendBtn.addEventListener('click', sendMessage);
    messageInput.addEventListener('keydown', (e) => {
        if (e.key === 'Enter' && !e.shiftKey) {
            e.preventDefault();
            sendMessage();
        }
    });
    messageInput.addEventListener('input', () => {
        autoResizeTextarea();
        sendBtn.disabled = !messageInput.value.trim();
    });

    // Nav items -> panels
    gmailNavItem.addEventListener('click', () => openPanel('gmail'));
    calendarNavItem.addEventListener('click', () => openPanel('calendar'));
    gchatNavItem.addEventListener('click', () => openPanel('gchat'));
    githubNavItem.addEventListener('click', () => openPanel('github'));

    // Close panel buttons
    document.querySelectorAll('.close-panel-btn').forEach(btn => {
        btn.addEventListener('click', () => closeAllPanels());
    });

    // Start chat buttons
    document.querySelectorAll('.start-chat-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            closeAllPanels();
            messageInput.focus();
        });
    });

    // Capabilities modal
    capabilitiesNavItem.addEventListener('click', async () => {
        capabilitiesModal.style.display = 'flex';
        await loadCapabilities();
    });
    closeCapabilitiesBtn.addEventListener('click', () => {
        capabilitiesModal.style.display = 'none';
    });
    capabilitiesModal.addEventListener('click', (e) => {
        if (e.target === capabilitiesModal) capabilitiesModal.style.display = 'none';
    });

    // Service tabs in modal
    document.getElementById('serviceTabs').addEventListener('click', (e) => {
        const tab = e.target.closest('.service-tab');
        if (!tab) return;
        document.querySelectorAll('.service-tab').forEach(t => t.classList.remove('active'));
        tab.classList.add('active');
        activeFilter = tab.dataset.filter;
        loadCapabilities();
    });

    // Auth buttons
    authenticateBtn.addEventListener('click', initiateGoogleAuth);
    calendarAuthBtn.addEventListener('click', initiateCalendarAuth);
    gchatAuthBtn.addEventListener('click', initiateGchatAuth);
    githubAuthBtn.addEventListener('click', initiateGithubAuth);
    githubDisconnectBtn.addEventListener('click', disconnectGitHub);
    gmailReauthBtn.addEventListener('click', initiateGoogleAuth);
    calendarReauthBtn.addEventListener('click', initiateCalendarAuth);
    gchatReauthBtn.addEventListener('click', initiateGchatAuth);
    githubReauthBtn.addEventListener('click', initiateGithubAuth);

    // Quick action buttons (all panels)
    document.querySelectorAll('.quick-action-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            const prompt = btn.dataset.prompt;
            closeAllPanels();
            messageInput.value = prompt;
            autoResizeTextarea();
            sendBtn.disabled = false;
            messageInput.focus();
        });
    });

    // Quick nav buttons in sidebar
    document.querySelectorAll('.quick-nav').forEach(btn => {
        btn.addEventListener('click', () => {
            const prompt = btn.dataset.prompt;
            messageInput.value = prompt;
            autoResizeTextarea();
            sendBtn.disabled = false;
            sendMessage();
        });
    });
}

// Panel management
function openPanel(service) {
    closeAllPanels();
    const panels = { gmail: gmailPanel, calendar: calendarPanel, gchat: gchatPanel, github: githubPanel };
    const navItems = { gmail: gmailNavItem, calendar: calendarNavItem, gchat: gchatNavItem, github: githubNavItem };
    if (panels[service]) {
        panels[service].classList.add('active');
        navItems[service].classList.add('active');
    }
}

function closeAllPanels() {
    [gmailPanel, calendarPanel, gchatPanel, githubPanel].forEach(p => p.classList.remove('active'));
    [gmailNavItem, calendarNavItem, gchatNavItem, githubNavItem].forEach(n => n.classList.remove('active'));
}

// Load capabilities into the modal dynamically
async function loadCapabilities() {
    const container = document.getElementById('toolsGrid');
    container.innerHTML = '<div style="color:var(--text-secondary)">Loading tools from server...</div>';

    try {
        const response = await fetch('/api/tools');
        const data = await response.json();
        const { services, totalTools } = data;

        toolCountBadge.textContent = `${totalTools} tools`;
        modalTitle.textContent = `All ${totalTools} Tools`;

        // Update status text
        const connectedNames = services.filter(s => s.connected).map(s => s.service);
        toolStatusText.textContent = connectedNames.length > 0 ? connectedNames.join(', ') + ' connected' : 'No services connected';
        toolStatusText.style.color = connectedNames.length > 0 ? 'var(--success)' : 'var(--error)';

        container.innerHTML = '';

        const serviceConfig = {
            gmail: { label: 'Gmail', dot: 'gmail', categories: GMAIL_CATEGORIES },
            calendar: { label: 'Google Calendar', dot: 'calendar', categories: CALENDAR_CATEGORIES },
            gchat: { label: 'Google Chat', dot: 'gchat', categories: GCHAT_CATEGORIES },
            github: { label: 'GitHub', dot: 'github', categories: GITHUB_CATEGORIES }
        };

        for (const svc of services) {
            if (activeFilter !== 'all' && svc.service !== activeFilter) continue;

            const config = serviceConfig[svc.service];
            if (!config) continue;

            // Service header
            const svcHeader = document.createElement('div');
            svcHeader.className = 'service-section-header';
            svcHeader.innerHTML = `<span class="service-dot ${config.dot}"></span> ${config.label} (${svc.tools.length}) ${svc.connected ? '<span style="color:var(--success);font-size:0.75rem">Connected</span>' : '<span style="color:var(--text-muted);font-size:0.75rem">Not connected</span>'}`;
            container.appendChild(svcHeader);

            // Categories
            for (const [category, toolNames] of Object.entries(config.categories)) {
                const categoryTools = svc.tools.filter(t => toolNames.includes(t.function.name));
                if (categoryTools.length === 0) continue;

                const header = document.createElement('div');
                header.className = 'tool-category-header';
                header.textContent = `${category} (${categoryTools.length})`;
                container.appendChild(header);

                const grid = document.createElement('div');
                grid.className = 'tools-grid';
                categoryTools.forEach(t => {
                    const name = t.function.name;
                    const desc = t.function.description;
                    const icon = TOOL_ICONS[name] || '&#128295;';
                    const card = document.createElement('div');
                    card.className = 'tool-card';
                    card.innerHTML = `
                        <div class="tool-icon">${icon}</div>
                        <div class="tool-info">
                            <h4>${name.replace(/_/g, ' ')}</h4>
                            <p>${desc}</p>
                            <code>${name}</code>
                        </div>
                    `;
                    grid.appendChild(card);
                });
                container.appendChild(grid);
            }
        }

    } catch (error) {
        console.error('Failed to load tools:', error);
        container.innerHTML = '<div style="color:var(--error)">Failed to load capabilities</div>';
    }
}

// Auto-resize textarea
function autoResizeTextarea() {
    messageInput.style.height = 'auto';
    messageInput.style.height = Math.min(messageInput.scrollHeight, 150) + 'px';
}

// Check all service statuses
async function checkAllStatuses() {
    checkGmailStatus();
    checkCalendarStatus();
    checkGchatStatus();
    checkGitHubStatus();

    setInterval(() => {
        checkGmailStatus();
        checkCalendarStatus();
        checkGchatStatus();
        checkGitHubStatus();
    }, 5000);
}

// Gmail status
async function checkGmailStatus() {
    try {
        const response = await fetch('/api/gmail/status');
        const data = await response.json();
        updateGmailStatus(data);
    } catch (error) {
        updateGmailStatus({ authenticated: false, credentialsConfigured: false });
    }
}

function updateGmailStatus(data) {
    const statusDot = gmailStatus.querySelector('.status-dot');
    isGmailConnected = data.authenticated;

    if (data.authenticated) {
        statusDot.className = 'status-dot connected';
        authSection.style.display = 'none';
        connectedSection.style.display = 'block';
        setupSection.style.display = 'none';
        gmailBadge.style.display = 'inline-flex';
    } else if (!data.credentialsConfigured) {
        statusDot.className = 'status-dot disconnected';
        authSection.style.display = 'none';
        connectedSection.style.display = 'none';
        setupSection.style.display = 'block';
        gmailBadge.style.display = 'none';
    } else {
        statusDot.className = 'status-dot disconnected';
        authSection.style.display = 'block';
        connectedSection.style.display = 'none';
        setupSection.style.display = 'none';
        gmailBadge.style.display = 'none';
    }
}

// Calendar status
async function checkCalendarStatus() {
    try {
        const response = await fetch('/api/calendar/status');
        const data = await response.json();
        updateCalendarStatus(data);
    } catch (error) {
        updateCalendarStatus({ authenticated: false });
    }
}

function updateCalendarStatus(data) {
    const statusDot = calendarStatus.querySelector('.status-dot');
    isCalendarConnected = data.authenticated;

    if (data.authenticated) {
        statusDot.className = 'status-dot connected';
        calendarAuthSection.style.display = 'none';
        calendarConnectedSection.style.display = 'block';
        calendarBadge.style.display = 'inline-flex';
    } else {
        statusDot.className = 'status-dot disconnected';
        calendarAuthSection.style.display = 'block';
        calendarConnectedSection.style.display = 'none';
        calendarBadge.style.display = 'none';
    }
}

// GitHub status
async function checkGitHubStatus() {
    try {
        const response = await fetch('/api/github/status');
        const data = await response.json();
        updateGitHubStatus(data);
    } catch (error) {
        updateGitHubStatus({ authenticated: false });
    }
}

function updateGitHubStatus(data) {
    const statusDot = githubStatus.querySelector('.status-dot');
    isGithubConnected = data.authenticated;
    const usernameText = data.username ? `Connected as @${data.username}. 20 tools ready.` : '20 tools ready.';
    const methodText = data.authMethod === 'oauth' ? 'OAuth' : (data.authMethod === 'pat' ? 'PAT' : 'OAuth');

    document.getElementById('githubUsername').textContent = `${usernameText} (${methodText})`;

    if (!data.oauthConfigured && githubAuthNote) {
        githubAuthNote.textContent = 'GitHub OAuth is not configured on the server. Add GITHUB_CLIENT_ID and GITHUB_CLIENT_SECRET to .env.';
    } else if (githubAuthNote) {
        githubAuthNote.textContent = 'Uses secure GitHub OAuth. You can reauthenticate later to switch accounts.';
    }

    if (data.authenticated) {
        statusDot.className = 'status-dot connected';
        githubAuthSection.style.display = 'none';
        githubConnectedSection.style.display = 'block';
        githubBadge.style.display = 'inline-flex';
    } else {
        statusDot.className = 'status-dot disconnected';
        githubAuthSection.style.display = 'block';
        githubConnectedSection.style.display = 'none';
        githubBadge.style.display = 'none';
    }
}

// Google Chat status
async function checkGchatStatus() {
    try {
        const response = await fetch('/api/gchat/status');
        const data = await response.json();
        updateGchatStatus(data);
    } catch (error) {
        updateGchatStatus({ authenticated: false });
    }
}

function updateGchatStatus(data) {
    const statusDot = gchatStatus.querySelector('.status-dot');
    isGchatConnected = data.authenticated;

    if (data.authenticated) {
        statusDot.className = 'status-dot connected';
        gchatAuthSection.style.display = 'none';
        gchatConnectedSection.style.display = 'block';
        gchatBadge.style.display = 'inline-flex';
    } else {
        statusDot.className = 'status-dot disconnected';
        gchatAuthSection.style.display = 'block';
        gchatConnectedSection.style.display = 'none';
        gchatBadge.style.display = 'none';
    }
}

// Google OAuth (Gmail + Calendar)
async function initiateGoogleAuth() {
    try {
        authenticateBtn.disabled = true;
        authenticateBtn.innerHTML = `<svg class="spinner" width="20" height="20" viewBox="0 0 24 24"><circle cx="12" cy="12" r="10" stroke="currentColor" stroke-width="3" fill="none" stroke-dasharray="30 70" /></svg> Connecting...`;

        const response = await fetch('/api/gmail/auth');
        const data = await response.json();

        if (data.authUrl) {
            const popup = window.open(data.authUrl, 'Google Authentication', 'width=600,height=700,left=200,top=100');
            const checkClosed = setInterval(() => {
                if (popup.closed) {
                    clearInterval(checkClosed);
                    checkGmailStatus();
                    checkCalendarStatus();
                    checkGchatStatus();
                    resetGoogleAuthButton();
                }
            }, 500);
        } else if (data.setupRequired) {
            alert('Please set up Google Cloud credentials first.');
            resetGoogleAuthButton();
        }
    } catch (error) {
        console.error('Google auth error:', error);
        alert('Failed to initiate authentication');
        resetGoogleAuthButton();
    }
}

async function initiateCalendarAuth() {
    try {
        calendarAuthBtn.disabled = true;
        calendarAuthBtn.innerHTML = `<svg class="spinner" width="20" height="20" viewBox="0 0 24 24"><circle cx="12" cy="12" r="10" stroke="currentColor" stroke-width="3" fill="none" stroke-dasharray="30 70" /></svg> Connecting...`;

        const response = await fetch('/api/calendar/connect');
        const data = await response.json();

        if (data.authUrl) {
            const popup = window.open(data.authUrl, 'Google Calendar Auth', 'width=600,height=700,left=200,top=100');
            const checkClosed = setInterval(() => {
                if (popup.closed) {
                    clearInterval(checkClosed);
                    checkGmailStatus();
                    checkCalendarStatus();
                    checkGchatStatus();
                    resetCalendarAuthButton();
                }
            }, 500);
        } else if (data.setupRequired) {
            alert('Please set up Google Cloud credentials first.');
            resetCalendarAuthButton();
        }
    } catch (error) {
        console.error('Calendar auth error:', error);
        alert('Failed to initiate calendar authentication');
        resetCalendarAuthButton();
    }
}

function resetGoogleAuthButton() {
    authenticateBtn.disabled = false;
    authenticateBtn.innerHTML = `<svg viewBox="0 0 24 24" width="20" height="20"><path fill="#fff" d="M22.56 12.25c0-.78-.07-1.53-.2-2.25H12v4.26h5.92c-.26 1.37-1.04 2.53-2.21 3.31v2.77h3.57c2.08-1.92 3.28-4.74 3.28-8.09z"/><path fill="#fff" d="M12 23c2.97 0 5.46-.98 7.28-2.66l-3.57-2.77c-.98.66-2.23 1.06-3.71 1.06-2.86 0-5.29-1.93-6.16-4.53H2.18v2.84C3.99 20.53 7.7 23 12 23z"/><path fill="#fff" d="M5.84 14.09c-.22-.66-.35-1.36-.35-2.09s.13-1.43.35-2.09V7.07H2.18C1.43 8.55 1 10.22 1 12s.43 3.45 1.18 4.93l2.85-2.22.81-.62z"/><path fill="#fff" d="M12 5.38c1.62 0 3.06.56 4.21 1.64l3.15-3.15C17.45 2.09 14.97 1 12 1 7.7 1 3.99 3.47 2.18 7.07l3.66 2.84c.87-2.6 3.3-4.53 6.16-4.53z"/></svg> Sign in with Google`;
}

function resetCalendarAuthButton() {
    calendarAuthBtn.disabled = false;
    calendarAuthBtn.innerHTML = `<svg viewBox="0 0 24 24" width="20" height="20"><path fill="#fff" d="M22.56 12.25c0-.78-.07-1.53-.2-2.25H12v4.26h5.92c-.26 1.37-1.04 2.53-2.21 3.31v2.77h3.57c2.08-1.92 3.28-4.74 3.28-8.09z"/><path fill="#fff" d="M12 23c2.97 0 5.46-.98 7.28-2.66l-3.57-2.77c-.98.66-2.23 1.06-3.71 1.06-2.86 0-5.29-1.93-6.16-4.53H2.18v2.84C3.99 20.53 7.7 23 12 23z"/><path fill="#fff" d="M5.84 14.09c-.22-.66-.35-1.36-.35-2.09s.13-1.43.35-2.09V7.07H2.18C1.43 8.55 1 10.22 1 12s.43 3.45 1.18 4.93l2.85-2.22.81-.62z"/><path fill="#fff" d="M12 5.38c1.62 0 3.06.56 4.21 1.64l3.15-3.15C17.45 2.09 14.97 1 12 1 7.7 1 3.99 3.47 2.18 7.07l3.66 2.84c.87-2.6 3.3-4.53 6.16-4.53z"/></svg> Sign in with Google`;
}

async function initiateGchatAuth() {
    try {
        gchatAuthBtn.disabled = true;
        gchatAuthBtn.innerHTML = `<svg class="spinner" width="20" height="20" viewBox="0 0 24 24"><circle cx="12" cy="12" r="10" stroke="currentColor" stroke-width="3" fill="none" stroke-dasharray="30 70" /></svg> Connecting...`;

        const response = await fetch('/api/gchat/connect');
        const contentType = (response.headers.get('content-type') || '').toLowerCase();
        let data = {};
        if (contentType.includes('application/json')) {
            data = await response.json();
        } else {
            const text = await response.text();
            if (response.status === 404) {
                throw new Error('Google Chat auth route not found on backend. Restart the server so latest routes are loaded.');
            }
            throw new Error(`Unexpected non-JSON response from backend (${response.status}): ${text.slice(0, 120)}`);
        }

        if (!response.ok) {
            throw new Error(data.error || `Google Chat auth failed with status ${response.status}`);
        }

        if (data.authUrl) {
            const popup = window.open(data.authUrl, 'Google Chat Auth', 'width=600,height=700,left=200,top=100');
            const checkClosed = setInterval(() => {
                if (popup.closed) {
                    clearInterval(checkClosed);
                    checkGmailStatus();
                    checkCalendarStatus();
                    checkGchatStatus();
                    resetGchatAuthButton();
                }
            }, 500);
        } else if (data.setupRequired) {
            alert('Please set up Google Cloud credentials first.');
            resetGchatAuthButton();
        } else {
            alert(data.error || 'Failed to initiate Google Chat authentication');
            resetGchatAuthButton();
        }
    } catch (error) {
        console.error('Google Chat auth error:', error);
        alert(error.message || 'Failed to initiate Google Chat authentication');
        resetGchatAuthButton();
    }
}

function resetGchatAuthButton() {
    gchatAuthBtn.disabled = false;
    gchatAuthBtn.innerHTML = `<svg viewBox="0 0 24 24" width="20" height="20"><path fill="#fff" d="M22.56 12.25c0-.78-.07-1.53-.2-2.25H12v4.26h5.92c-.26 1.37-1.04 2.53-2.21 3.31v2.77h3.57c2.08-1.92 3.28-4.74 3.28-8.09z"/><path fill="#fff" d="M12 23c2.97 0 5.46-.98 7.28-2.66l-3.57-2.77c-.98.66-2.23 1.06-3.71 1.06-2.86 0-5.29-1.93-6.16-4.53H2.18v2.84C3.99 20.53 7.7 23 12 23z"/><path fill="#fff" d="M5.84 14.09c-.22-.66-.35-1.36-.35-2.09s.13-1.43.35-2.09V7.07H2.18C1.43 8.55 1 10.22 1 12s.43 3.45 1.18 4.93l2.85-2.22.81-.62z"/><path fill="#fff" d="M12 5.38c1.62 0 3.06.56 4.21 1.64l3.15-3.15C17.45 2.09 14.97 1 12 1 7.7 1 3.99 3.47 2.18 7.07l3.66 2.84c.87-2.6 3.3-4.53 6.16-4.53z"/></svg> Sign in with Google`;
}

// GitHub OAuth connect
async function initiateGithubAuth() {
    try {
        githubAuthBtn.disabled = true;
        githubAuthBtn.innerHTML = `<svg class="spinner" width="20" height="20" viewBox="0 0 24 24"><circle cx="12" cy="12" r="10" stroke="currentColor" stroke-width="3" fill="none" stroke-dasharray="30 70" /></svg> Connecting...`;

        const response = await fetch('/api/github/auth');
        const contentType = (response.headers.get('content-type') || '').toLowerCase();
        let data = {};
        if (contentType.includes('application/json')) {
            data = await response.json();
        } else {
            const text = await response.text();
            if (response.status === 404) {
                throw new Error('GitHub auth route not found on backend. Restart the server so latest routes are loaded.');
            }
            throw new Error(`Unexpected non-JSON response from backend (${response.status}): ${text.slice(0, 120)}`);
        }

        if (!response.ok) {
            throw new Error(data.error || `GitHub auth failed with status ${response.status}`);
        }

        if (data.authUrl) {
            const popup = window.open(data.authUrl, 'GitHub Authentication', 'width=600,height=700,left=200,top=100');
            const checkClosed = setInterval(() => {
                if (popup.closed) {
                    clearInterval(checkClosed);
                    checkGitHubStatus();
                    resetGithubAuthButton();
                }
            }, 500);
        } else if (data.setupRequired) {
            alert('Please configure GITHUB_CLIENT_ID and GITHUB_CLIENT_SECRET in .env first.');
            resetGithubAuthButton();
        } else {
            alert(data.error || 'Failed to start GitHub authentication');
            resetGithubAuthButton();
        }
    } catch (error) {
        console.error('GitHub OAuth error:', error);
        alert(error.message || 'Failed to initiate GitHub authentication');
        resetGithubAuthButton();
    }
}

function resetGithubAuthButton() {
    githubAuthBtn.disabled = false;
    githubAuthBtn.innerHTML = `<svg viewBox="0 0 24 24" width="20" height="20"><path fill="#fff" d="M12 0C5.37 0 0 5.37 0 12c0 5.31 3.435 9.795 8.205 11.385.6.105.825-.255.825-.57 0-.285-.015-1.23-.015-2.235-3.015.555-3.795-.735-4.035-1.41-.135-.345-.72-1.41-1.23-1.695-.42-.225-1.02-.78-.015-.795.945-.015 1.62.87 1.845 1.23 1.08 1.815 2.805 1.305 3.495.99.105-.78.42-1.305.765-1.605-2.67-.3-5.46-1.335-5.46-5.925 0-1.305.465-2.385 1.23-3.225-.12-.3-.54-1.53.12-3.18 0 0 1.005-.315 3.3 1.23.96-.27 1.98-.405 3-.405s2.04.135 3 .405c2.295-1.56 3.3-1.23 3.3-1.23.66 1.65.24 2.88.12 3.18.765.84 1.23 1.905 1.23 3.225 0 4.605-2.805 5.625-5.475 5.925.435.375.81 1.095.81 2.22 0 1.605-.015 2.895-.015 3.3 0 .315.225.69.825.57A12.02 12.02 0 0 0 24 12c0-6.63-5.37-12-12-12z"/></svg> Sign in with GitHub`;
}

async function disconnectGitHub() {
    try {
        await fetch('/api/github/disconnect', { method: 'POST' });
        checkGitHubStatus();
    } catch (error) {
        console.error('GitHub disconnect error:', error);
    }
}

// Send message
async function sendMessage() {
    const message = messageInput.value.trim();
    if (!message) return;

    messageInput.value = '';
    autoResizeTextarea();
    sendBtn.disabled = true;

    addMessage('user', message);
    const typingId = addTypingIndicator();

    turnsBadge.style.display = 'inline-flex';
    turnsCount.textContent = '...';

    try {
        const response = await fetch('/api/chat', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ message, history: chatHistory })
        });

        const data = await response.json();
        removeTypingIndicator(typingId);

        if (data.error) {
            addMessage('assistant', `<p style="color: #ef4444;">Error: ${escapeHtml(data.error)}</p>`);
            turnsBadge.style.display = 'none';
        } else {
            if (data.turnsUsed > 0) {
                turnsCount.textContent = data.turnsUsed;
            } else {
                turnsBadge.style.display = 'none';
            }

            let responseHtml = '';

            if (data.steps && data.steps.length > 0) {
                responseHtml += formatStepsPipeline(data.steps);
            }

            responseHtml += formatResponse(data.response);

            if (data.toolResults && data.toolResults.length > 0) {
                responseHtml += formatToolResults(data.toolResults);
            }

            addMessage('assistant', responseHtml);

            chatHistory.push({ role: 'user', content: message });
            chatHistory.push({ role: 'assistant', content: data.response });

            if (chatHistory.length > 30) {
                chatHistory = chatHistory.slice(-30);
            }
        }
    } catch (error) {
        removeTypingIndicator(typingId);
        addMessage('assistant', `<p style="color: #ef4444;">Failed to send message. Please try again.</p>`);
        turnsBadge.style.display = 'none';
        console.error('Chat error:', error);
    }
}

// Format the step-by-step execution pipeline
function formatStepsPipeline(steps) {
    if (!steps || steps.length === 0) return '';

    const stepsHtml = steps.map((step, i) => {
        const icon = TOOL_ICONS[step.tool] || '&#128295;';
        const statusClass = step.success ? 'step-success' : 'step-error';
        const statusIcon = step.success ? '&#10003;' : '&#10007;';
        const toolName = step.tool.replace(/_/g, ' ');

        return `
            <div class="pipeline-step ${statusClass}">
                <div class="step-number">${i + 1}</div>
                <div class="step-icon">${icon}</div>
                <div class="step-info">
                    <span class="step-name">${toolName}</span>
                    <span class="step-status">${statusIcon}</span>
                </div>
            </div>
            ${i < steps.length - 1 ? '<div class="pipeline-connector"></div>' : ''}
        `;
    }).join('');

    return `
        <div class="execution-pipeline">
            <div class="pipeline-header">Agent Execution (${steps.length} step${steps.length > 1 ? 's' : ''})</div>
            <div class="pipeline-steps">${stepsHtml}</div>
        </div>
    `;
}

// Format response text to HTML
function formatResponse(text) {
    if (!text) return '';

    return text
        .split('\n')
        .map(line => {
            if (line.trim().startsWith('- ') || line.trim().startsWith('* ')) {
                return `<li>${escapeHtml(line.trim().slice(2))}</li>`;
            }
            if (/^\d+\.\s/.test(line.trim())) {
                return `<li>${escapeHtml(line.trim().replace(/^\d+\.\s/, ''))}</li>`;
            }
            line = escapeHtml(line);
            if (line.includes('`')) {
                line = line.replace(/`([^`]+)`/g, '<code>$1</code>');
            }
            if (line.includes('**')) {
                line = line.replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>');
            }
            return line ? `<p>${line}</p>` : '';
        })
        .join('');
}

// Format tool results
function formatToolResults(results) {
    return results.map(result => {
        const isError = result.error;
        const icon = isError ? '&#10007;' : '&#10003;';
        const title = result.tool.replace(/_/g, ' ').replace(/\b\w/g, l => l.toUpperCase());
        const toolIcon = TOOL_ICONS[result.tool] || '&#128295;';

        let content = '';
        if (result.result) {
            // Email list results
            if (result.result.results && Array.isArray(result.result.results)) {
                content = result.result.results.map(email => `
                    <div class="email-card">
                        <div class="email-card-header">
                            <span class="email-card-subject">${escapeHtml(email.subject || '(no subject)')}</span>
                            <span class="email-card-date">${formatDate(email.date)}</span>
                        </div>
                        <div class="email-card-from">${escapeHtml(email.from || '')}</div>
                        ${email.snippet ? `<div class="email-card-snippet">${escapeHtml(email.snippet.slice(0, 120))}...</div>` : ''}
                    </div>
                `).join('');
            // Email body
            } else if (result.result.body) {
                content = `
                    <div class="email-card">
                        <div class="email-card-header">
                            <span class="email-card-subject">${escapeHtml(result.result.subject || '')}</span>
                        </div>
                        <div class="email-card-from">From: ${escapeHtml(result.result.from || '')}</div>
                        ${result.result.hasAttachments ? `<div class="email-card-attachments">${result.result.attachments.length} attachment(s)</div>` : ''}
                        <p style="margin-top: 0.5rem; white-space: pre-wrap; font-size:0.85rem; color:var(--text-secondary)">${escapeHtml((result.result.body || '').slice(0, 500))}${(result.result.body || '').length > 500 ? '...' : ''}</p>
                    </div>
                `;
            // Labels
            } else if (result.result.labels && Array.isArray(result.result.labels)) {
                content = `<div class="labels-list">${result.result.labels.map(l => `<span class="label-chip">${escapeHtml(l.name)}</span>`).join('')}</div>`;
            // Drafts
            } else if (result.result.drafts && Array.isArray(result.result.drafts)) {
                content = result.result.drafts.map(d => `
                    <div class="email-card">
                        <div class="email-card-header">
                            <span class="email-card-subject">${escapeHtml(d.subject || '(no subject)')}</span>
                        </div>
                        <div class="email-card-from">To: ${escapeHtml(d.to || 'Not set')}</div>
                    </div>
                `).join('');
            // Google Chat spaces
            } else if (result.result.spaces && Array.isArray(result.result.spaces)) {
                content = result.result.spaces.map(s => `
                    <div class="email-card">
                        <div class="email-card-header">
                            <span class="email-card-subject">${escapeHtml(s.displayName || s.name)}</span>
                            <span class="email-card-date">${escapeHtml(s.spaceType || '')}</span>
                        </div>
                        <div class="email-card-from"><code>${escapeHtml(s.name || '')}</code></div>
                    </div>
                `).join('');
            // Google Chat messages
            } else if (result.result.space && result.result.messages && Array.isArray(result.result.messages)) {
                content = result.result.messages.map(msg => `
                    <div class="email-card thread-msg">
                        <div class="email-card-header">
                            <span class="email-card-from" style="font-weight:600">${escapeHtml(msg.sender || '')}</span>
                            <span class="email-card-date">${formatDate(msg.createTime)}</span>
                        </div>
                        <div class="email-card-snippet">${escapeHtml(msg.text || '')}</div>
                    </div>
                `).join('');
            // Thread messages
            } else if (result.result.messages && Array.isArray(result.result.messages)) {
                content = result.result.messages.map(msg => `
                    <div class="email-card thread-msg">
                        <div class="email-card-header">
                            <span class="email-card-from" style="font-weight:600">${escapeHtml(msg.from || '')}</span>
                            <span class="email-card-date">${formatDate(msg.date)}</span>
                        </div>
                        <p style="margin-top:0.25rem;font-size:0.85rem;color:var(--text-secondary)">${escapeHtml((msg.body || msg.snippet || '').slice(0, 200))}...</p>
                    </div>
                `).join('');
            // Attachments
            } else if (result.result.attachments && Array.isArray(result.result.attachments)) {
                content = result.result.attachments.map(a => `
                    <div class="attachment-card">
                        <span class="attachment-icon">&#128206;</span>
                        <div>
                            <div style="font-weight:500">${escapeHtml(a.filename)}</div>
                            <div style="font-size:0.8rem;color:var(--text-muted)">${escapeHtml(a.mimeType)} &middot; ${formatSize(a.size)}</div>
                        </div>
                    </div>
                `).join('');
            // Calendar events
            } else if (result.result.events && Array.isArray(result.result.events)) {
                content = result.result.events.map(e => `
                    <div class="email-card">
                        <div class="email-card-header">
                            <span class="email-card-subject">${escapeHtml(e.summary || '(no title)')}</span>
                            <span class="email-card-date">${formatDate(e.start)}</span>
                        </div>
                        ${e.location ? `<div class="email-card-from">&#128205; ${escapeHtml(e.location)}</div>` : ''}
                        ${e.attendees && e.attendees.length > 0 ? `<div class="email-card-snippet">${e.attendees.length} attendee(s)</div>` : ''}
                        ${e.meetLink ? `<div class="email-card-snippet"><a href="${escapeHtml(e.meetLink)}" target="_blank" rel="noopener noreferrer">Open Meet Link</a></div>` : ''}
                    </div>
                `).join('');
            // Calendars list
            } else if (result.result.calendars && Array.isArray(result.result.calendars)) {
                content = result.result.calendars.map(c => `
                    <div class="email-card">
                        <div class="email-card-header">
                            <span class="email-card-subject">${escapeHtml(c.summary || c.id)}</span>
                            ${c.primary ? '<span style="color:var(--accent-primary);font-size:0.75rem">Primary</span>' : ''}
                        </div>
                        ${c.description ? `<div class="email-card-snippet">${escapeHtml(c.description)}</div>` : ''}
                    </div>
                `).join('');
            // GitHub repos
            } else if (result.result.repos && Array.isArray(result.result.repos)) {
                content = result.result.repos.map(r => `
                    <div class="email-card">
                        <div class="email-card-header">
                            <span class="email-card-subject">${escapeHtml(r.full_name || r.name)}</span>
                            <span class="email-card-date">${r.language || ''}</span>
                        </div>
                        ${r.description ? `<div class="email-card-snippet">${escapeHtml(r.description.slice(0, 100))}</div>` : ''}
                        <div class="email-card-from" style="margin-top:0.25rem">&#11088; ${r.stars || 0} &middot; &#128204; ${r.forks || 0}${r.private ? ' &middot; Private' : ''}</div>
                    </div>
                `).join('');
            // GitHub issues
            } else if (result.result.issues && Array.isArray(result.result.issues)) {
                content = result.result.issues.map(i => `
                    <div class="email-card">
                        <div class="email-card-header">
                            <span class="email-card-subject">#${i.number} ${escapeHtml(i.title)}</span>
                            <span class="email-card-date">${i.state}</span>
                        </div>
                        <div class="email-card-from">by ${escapeHtml(i.user || '')} &middot; ${i.comments || 0} comments</div>
                        ${i.labels && i.labels.length > 0 ? `<div class="labels-list" style="margin-top:0.25rem">${i.labels.map(l => `<span class="label-chip">${escapeHtml(l)}</span>`).join('')}</div>` : ''}
                    </div>
                `).join('');
            // GitHub PRs
            } else if (result.result.pullRequests && Array.isArray(result.result.pullRequests)) {
                content = result.result.pullRequests.map(p => `
                    <div class="email-card">
                        <div class="email-card-header">
                            <span class="email-card-subject">#${p.number} ${escapeHtml(p.title)}</span>
                            <span class="email-card-date">${p.state}${p.draft ? ' (draft)' : ''}</span>
                        </div>
                        <div class="email-card-from">${escapeHtml(p.head)} &#8594; ${escapeHtml(p.base)} &middot; by ${escapeHtml(p.user || '')}</div>
                    </div>
                `).join('');
            // GitHub branches
            } else if (result.result.branches && Array.isArray(result.result.branches)) {
                content = `<div class="labels-list">${result.result.branches.map(b => `<span class="label-chip">${escapeHtml(b.name)}${b.protected ? ' &#128274;' : ''}</span>`).join('')}</div>`;
            // GitHub commits
            } else if (result.result.commits && Array.isArray(result.result.commits)) {
                content = result.result.commits.map(c => `
                    <div class="email-card">
                        <div class="email-card-header">
                            <code style="font-size:0.8rem;color:var(--accent-secondary)">${escapeHtml(c.sha)}</code>
                            <span class="email-card-date">${formatDate(c.date)}</span>
                        </div>
                        <div class="email-card-from">${escapeHtml((c.message || '').split('\n')[0])}</div>
                        <div class="email-card-snippet">${escapeHtml(c.author || '')}</div>
                    </div>
                `).join('');
            // GitHub notifications
            } else if (result.result.notifications && Array.isArray(result.result.notifications)) {
                content = result.result.notifications.map(n => `
                    <div class="email-card">
                        <div class="email-card-header">
                            <span class="email-card-subject">${escapeHtml(n.subject.title)}</span>
                            <span class="email-card-date">${n.subject.type}</span>
                        </div>
                        <div class="email-card-from">${escapeHtml(n.repository)} &middot; ${n.reason}${n.unread ? ' &middot; Unread' : ''}</div>
                    </div>
                `).join('');
            // GitHub gists
            } else if (result.result.gists && Array.isArray(result.result.gists)) {
                content = result.result.gists.map(g => `
                    <div class="email-card">
                        <div class="email-card-header">
                            <span class="email-card-subject">${escapeHtml(g.description || g.files[0] || 'Untitled')}</span>
                            <span class="email-card-date">${g.public ? 'Public' : 'Secret'}</span>
                        </div>
                        <div class="email-card-from">${g.files.join(', ')}</div>
                    </div>
                `).join('');
            // Generic JSON display
            } else {
                const summary = result.result.message || JSON.stringify(result.result, null, 2);
                content = `<pre style="font-size:0.8rem;max-height:150px;overflow:auto">${escapeHtml(typeof summary === 'string' ? summary : JSON.stringify(summary, null, 2))}</pre>`;
            }
        } else if (result.error) {
            content = `<p style="color: #ef4444;">${escapeHtml(result.error)}</p>`;
        }

        return `
            <div class="tool-result ${isError ? 'tool-result-error' : ''}">
                <div class="tool-result-header">
                    <span class="tool-result-icon">${toolIcon}</span>
                    ${icon} ${title}
                </div>
                ${content}
            </div>
        `;
    }).join('');
}

// Add message to chat
function addMessage(role, content) {
    const messageDiv = document.createElement('div');
    messageDiv.className = `message ${role}`;

    const avatar = role === 'user'
        ? '<svg viewBox="0 0 24 24" width="20" height="20"><path fill="#fff" d="M12 12c2.21 0 4-1.79 4-4s-1.79-4-4-4-4 1.79-4 4 1.79 4 4 4zm0 2c-2.67 0-8 1.34-8 4v2h16v-2c0-2.66-5.33-4-8-4z"/></svg>'
        : '<svg viewBox="0 0 24 24" width="20" height="20"><circle cx="12" cy="12" r="10" fill="none" stroke="#6366f1" stroke-width="2"/><path fill="none" stroke="#6366f1" stroke-width="2" d="M8 12l3 3 5-6"/></svg>';

    messageDiv.innerHTML = `
        <div class="message-avatar">${avatar}</div>
        <div class="message-content">
            <div class="message-bubble">${content}</div>
        </div>
    `;

    chatMessages.appendChild(messageDiv);
    scrollToBottom();
}

// Add typing indicator
function addTypingIndicator() {
    const id = 'typing-' + Date.now();
    const typingDiv = document.createElement('div');
    typingDiv.id = id;
    typingDiv.className = 'message assistant';
    typingDiv.innerHTML = `
        <div class="message-avatar">
            <svg viewBox="0 0 24 24" width="20" height="20">
                <circle cx="12" cy="12" r="10" fill="none" stroke="#6366f1" stroke-width="2"/>
                <path fill="none" stroke="#6366f1" stroke-width="2" d="M8 12l3 3 5-6"/>
            </svg>
        </div>
        <div class="message-content">
            <div class="message-bubble">
                <div class="typing-indicator">
                    <span></span><span></span><span></span>
                    <span class="typing-text">Agent working...</span>
                </div>
            </div>
        </div>
    `;
    chatMessages.appendChild(typingDiv);
    scrollToBottom();
    return id;
}

// Remove typing indicator
function removeTypingIndicator(id) {
    const element = document.getElementById(id);
    if (element) element.remove();
}

// Scroll to bottom
function scrollToBottom() {
    chatMessages.scrollTop = chatMessages.scrollHeight;
}

// Escape HTML
function escapeHtml(text) {
    if (!text) return '';
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

// Format date
function formatDate(dateStr) {
    if (!dateStr) return '';
    try {
        const date = new Date(dateStr);
        const now = new Date();
        const diff = now - date;
        if (diff < 86400000) return date.toLocaleTimeString('en-US', { hour: 'numeric', minute: '2-digit' });
        if (diff < 604800000) return date.toLocaleDateString('en-US', { weekday: 'short' });
        return date.toLocaleDateString('en-US', { month: 'short', day: 'numeric' });
    } catch { return dateStr; }
}

// Format file size
function formatSize(bytes) {
    if (!bytes) return '0 B';
    const sizes = ['B', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(1024));
    return (bytes / Math.pow(1024, i)).toFixed(1) + ' ' + sizes[i];
}

// Spinner animation
const style = document.createElement('style');
style.textContent = `
    .spinner { animation: spin 1s linear infinite; }
    @keyframes spin { to { transform: rotate(360deg); } }
`;
document.head.appendChild(style);
