// DOM Elements
const chatMessages = document.getElementById('chatMessages');
const messageInput = document.getElementById('messageInput');
const sendBtn = document.getElementById('sendBtn');
const gmailNavItem = document.getElementById('gmailNavItem');
const calendarNavItem = document.getElementById('calendarNavItem');
const gchatNavItem = document.getElementById('gchatNavItem');
const driveNavItem = document.getElementById('driveNavItem');
const sheetsNavItem = document.getElementById('sheetsNavItem');
const githubNavItem = document.getElementById('githubNavItem');
const timerNavItem = document.getElementById('timerNavItem');
const gmailPanel = document.getElementById('gmailPanel');
const calendarPanel = document.getElementById('calendarPanel');
const gchatPanel = document.getElementById('gchatPanel');
const drivePanel = document.getElementById('drivePanel');
const sheetsPanel = document.getElementById('sheetsPanel');
const githubPanel = document.getElementById('githubPanel');
const timerPanel = document.getElementById('timerPanel');
const authenticateBtn = document.getElementById('authenticateBtn');
const calendarAuthBtn = document.getElementById('calendarAuthBtn');
const gchatAuthBtn = document.getElementById('gchatAuthBtn');
const driveAuthBtn = document.getElementById('driveAuthBtn');
const sheetsAuthBtn = document.getElementById('sheetsAuthBtn');
const githubAuthBtn = document.getElementById('githubAuthBtn');
const githubDisconnectBtn = document.getElementById('githubDisconnectBtn');
const gmailReauthBtn = document.getElementById('gmailReauthBtn');
const calendarReauthBtn = document.getElementById('calendarReauthBtn');
const gchatReauthBtn = document.getElementById('gchatReauthBtn');
const driveReauthBtn = document.getElementById('driveReauthBtn');
const sheetsReauthBtn = document.getElementById('sheetsReauthBtn');
const githubReauthBtn = document.getElementById('githubReauthBtn');
const docsReauthBtn = document.getElementById('docsReauthBtn');
const githubAuthNote = document.getElementById('githubAuthNote');
const authSection = document.getElementById('authSection');
const connectedSection = document.getElementById('connectedSection');
const setupSection = document.getElementById('setupSection');
const calendarAuthSection = document.getElementById('calendarAuthSection');
const calendarConnectedSection = document.getElementById('calendarConnectedSection');
const gchatAuthSection = document.getElementById('gchatAuthSection');
const gchatConnectedSection = document.getElementById('gchatConnectedSection');
const driveAuthSection = document.getElementById('driveAuthSection');
const driveConnectedSection = document.getElementById('driveConnectedSection');
const sheetsAuthSection = document.getElementById('sheetsAuthSection');
const sheetsConnectedSection = document.getElementById('sheetsConnectedSection');
const githubAuthSection = document.getElementById('githubAuthSection');
const githubConnectedSection = document.getElementById('githubConnectedSection');
const gmailStatus = document.getElementById('gmailStatus');
const calendarStatus = document.getElementById('calendarStatus');
const gchatStatus = document.getElementById('gchatStatus');
const driveStatus = document.getElementById('driveStatus');
const sheetsStatus = document.getElementById('sheetsStatus');
const githubStatus = document.getElementById('githubStatus');
const timerStatus = document.getElementById('timerStatus');
const gmailBadge = document.getElementById('gmailBadge');
const calendarBadge = document.getElementById('calendarBadge');
const gchatBadge = document.getElementById('gchatBadge');
const driveBadge = document.getElementById('driveBadge');
const sheetsBadge = document.getElementById('sheetsBadge');
const githubBadge = document.getElementById('githubBadge');
const timerBadge = document.getElementById('timerBadge');
const docsNavItem = document.getElementById('docsNavItem');
const docsPanel = document.getElementById('docsPanel');
const docsAuthSection = document.getElementById('docsAuthSection');
const docsConnectedSection = document.getElementById('docsConnectedSection');
const docsStatus = document.getElementById('docsStatus');
const docsBadge = document.getElementById('docsBadge');
const teamsNavItem = document.getElementById('teamsNavItem');
const teamsPanel = document.getElementById('teamsPanel');
const teamsAuthBtn = document.getElementById('teamsAuthBtn');
const teamsDisconnectBtn = document.getElementById('teamsDisconnectBtn');
const teamsAuthSection = document.getElementById('teamsAuthSection');
const teamsConnectedSection = document.getElementById('teamsConnectedSection');
const teamsSetupSection = document.getElementById('teamsSetupSection');
const teamsStatus = document.getElementById('teamsStatus');
const teamsBadge = document.getElementById('teamsBadge');
const teamsUserInfo = document.getElementById('teamsUserInfo');
const outlookNavItem = document.getElementById('outlookNavItem');
const outlookPanel = document.getElementById('outlookPanel');
const outlookAuthBtn = document.getElementById('outlookAuthBtn');
const outlookDisconnectBtn = document.getElementById('outlookDisconnectBtn');
const outlookReauthBtn = document.getElementById('outlookReauthBtn');
const outlookAuthNote = document.getElementById('outlookAuthNote');
const outlookAuthSection = document.getElementById('outlookAuthSection');
const outlookConnectedSection = document.getElementById('outlookConnectedSection');
const outlookStatus = document.getElementById('outlookStatus');
const outlookBadge = document.getElementById('outlookBadge');
const turnsBadge = document.getElementById('turnsBadge');
const turnsCount = document.getElementById('turnsCount');
const capabilitiesNavItem = document.getElementById('capabilitiesNavItem');
const capabilitiesModal = document.getElementById('capabilitiesModal');
const closeCapabilitiesBtn = document.getElementById('closeCapabilitiesBtn');
const toolCountBadge = document.getElementById('toolCountBadge');
const toolStatusText = document.getElementById('toolStatusText');
const modalTitle = document.getElementById('modalTitle');
const timerTaskNameInput = document.getElementById('timerTaskNameInput');
const timerTaskTimeInput = document.getElementById('timerTaskTimeInput');
const timerTaskInstructionInput = document.getElementById('timerTaskInstructionInput');
const timerTaskEnabledInput = document.getElementById('timerTaskEnabledInput');
const timerTaskSaveBtn = document.getElementById('timerTaskSaveBtn');
const timerTaskRefreshBtn = document.getElementById('timerTaskRefreshBtn');
const timerTaskList = document.getElementById('timerTaskList');
const timerTaskStatusText = document.getElementById('timerTaskStatusText');

// State
let chatHistory = [];
let isGmailConnected = false;
let isCalendarConnected = false;
let isGchatConnected = false;
let isDriveConnected = false;
let isSheetsConnected = false;
let isGithubConnected = false;
let isOutlookConnected = false;
let isDocsConnected = false;
let isTeamsConnected = false;
let isTimerConnected = false;
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
    // Google Drive (10)
    list_drive_files: '&#128193;', get_drive_file: '&#128196;', create_drive_folder: '&#128193;',
    create_drive_file: '&#10133;', update_drive_file: '&#9998;', delete_drive_file: '&#128465;',
    copy_drive_file: '&#128209;', move_drive_file: '&#10145;', share_drive_file: '&#128101;',
    download_drive_file: '&#128229;',
    // Google Sheets
    list_spreadsheets: '&#128202;', create_spreadsheet: '&#10133;', get_spreadsheet: '&#128196;',
    list_sheet_tabs: '&#128203;', add_sheet_tab: '&#10133;', delete_sheet_tab: '&#10134;',
    read_sheet_values: '&#128214;', update_sheet_values: '&#9998;', update_timesheet_hours: '&#9201;',
    append_sheet_values: '&#128228;',
    clear_sheet_values: '&#128465;',
    // GitHub (20)
    list_repos: '&#128193;', get_repo: '&#128196;', create_repo: '&#10133;',
    list_issues: '&#128196;', create_issue: '&#10133;', update_issue: '&#9998;',
    list_pull_requests: '&#128259;', get_pull_request: '&#128196;',
    create_pull_request: '&#10133;', merge_pull_request: '&#128279;',
    list_branches: '&#128204;', create_branch: '&#128204;',
    get_file_content: '&#128196;', create_or_update_file: '&#128196;',
    search_repos: '&#128269;', search_code: '&#128269;',
    list_commits: '&#128221;', get_user_profile: '&#128100;',
    list_notifications: '&#128276;', list_gists: '&#128221;',
    // Outlook (18)
    outlook_send_email: '&#9993;', outlook_list_emails: '&#128203;',
    outlook_read_email: '&#128214;', outlook_search_emails: '&#128269;',
    outlook_reply_to_email: '&#8617;', outlook_forward_email: '&#10145;',
    outlook_delete_email: '&#128465;', outlook_move_email: '&#10145;',
    outlook_mark_as_read: '&#9745;', outlook_mark_as_unread: '&#9746;',
    outlook_list_folders: '&#128193;', outlook_create_folder: '&#10133;',
    outlook_get_attachments: '&#128206;', outlook_create_draft: '&#128221;',
    outlook_send_draft: '&#128228;', outlook_list_drafts: '&#128196;',
    outlook_flag_email: '&#127988;', outlook_get_user_profile: '&#128100;',
    // Google Docs (8)
    list_documents: '&#128196;', get_document: '&#128196;', create_document: '&#10133;',
    insert_text: '&#9998;', replace_text: '&#128260;', delete_content: '&#128465;',
    append_text: '&#128228;', get_document_text: '&#128214;',
    // Microsoft Teams (10)
    teams_list_teams: '&#128101;', teams_get_team: '&#128196;',
    teams_list_channels: '&#128193;', teams_send_channel_message: '&#9993;',
    teams_list_channel_messages: '&#128221;', teams_list_chats: '&#128172;',
    teams_send_chat_message: '&#9993;', teams_list_chat_messages: '&#128221;',
    teams_create_chat: '&#10133;', teams_get_chat_members: '&#128101;'
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

const DRIVE_CATEGORIES = {
    'Browse': ['list_drive_files', 'get_drive_file', 'download_drive_file'],
    'Create & Edit': ['create_drive_folder', 'create_drive_file', 'update_drive_file'],
    'Manage': ['copy_drive_file', 'move_drive_file', 'share_drive_file', 'delete_drive_file']
};

const SHEETS_CATEGORIES = {
    'Discovery': ['list_spreadsheets', 'get_spreadsheet', 'list_sheet_tabs'],
    'Structure': ['create_spreadsheet', 'add_sheet_tab', 'delete_sheet_tab'],
    'Data': ['read_sheet_values', 'update_sheet_values', 'update_timesheet_hours', 'append_sheet_values', 'clear_sheet_values']
};

const GITHUB_CATEGORIES = {
    'Repositories': ['list_repos', 'get_repo', 'create_repo', 'search_repos'],
    'Issues': ['list_issues', 'create_issue', 'update_issue'],
    'Pull Requests': ['list_pull_requests', 'get_pull_request', 'create_pull_request', 'merge_pull_request'],
    'Code & Branches': ['list_branches', 'create_branch', 'get_file_content', 'create_or_update_file', 'search_code'],
    'Activity': ['list_commits', 'get_user_profile', 'list_notifications', 'list_gists']
};

const OUTLOOK_CATEGORIES = {
    'Core': ['outlook_send_email', 'outlook_search_emails', 'outlook_read_email', 'outlook_list_emails'],
    'Actions': ['outlook_reply_to_email', 'outlook_forward_email', 'outlook_delete_email', 'outlook_move_email'],
    'Status': ['outlook_mark_as_read', 'outlook_mark_as_unread', 'outlook_flag_email'],
    'Folders': ['outlook_list_folders', 'outlook_create_folder'],
    'Drafts': ['outlook_create_draft', 'outlook_list_drafts', 'outlook_send_draft'],
    'Advanced': ['outlook_get_attachments', 'outlook_get_user_profile']
};

const DOCS_CATEGORIES = {
    'Core': ['list_documents', 'get_document', 'create_document'],
    'Edit': ['insert_text', 'replace_text', 'delete_content', 'append_text'],
    'Read': ['get_document_text']
};

const TEAMS_CATEGORIES = {
    'Teams': ['teams_list_teams', 'teams_get_team'],
    'Channels': ['teams_list_channels', 'teams_send_channel_message', 'teams_list_channel_messages'],
    'Chats': ['teams_list_chats', 'teams_send_chat_message', 'teams_list_chat_messages', 'teams_create_chat', 'teams_get_chat_members']
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
    driveNavItem.addEventListener('click', () => openPanel('drive'));
    sheetsNavItem.addEventListener('click', () => openPanel('sheets'));
    githubNavItem.addEventListener('click', () => openPanel('github'));
    docsNavItem.addEventListener('click', () => openPanel('docs'));
    teamsNavItem.addEventListener('click', () => openPanel('teams'));
    outlookNavItem.addEventListener('click', () => openPanel('outlook'));
    timerNavItem.addEventListener('click', () => {
        openPanel('timer');
        loadTimerTasks();
    });

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
    driveAuthBtn.addEventListener('click', initiateDriveAuth);
    sheetsAuthBtn.addEventListener('click', initiateSheetsAuth);
    githubAuthBtn.addEventListener('click', initiateGithubAuth);
    githubDisconnectBtn.addEventListener('click', disconnectGitHub);
    gmailReauthBtn.addEventListener('click', initiateGoogleAuth);
    calendarReauthBtn.addEventListener('click', initiateCalendarAuth);
    gchatReauthBtn.addEventListener('click', initiateGchatAuth);
    driveReauthBtn.addEventListener('click', initiateDriveAuth);
    sheetsReauthBtn.addEventListener('click', initiateSheetsAuth);
    docsReauthBtn.addEventListener('click', initiateGoogleAuth);
    githubReauthBtn.addEventListener('click', initiateGithubAuth);
    teamsAuthBtn.addEventListener('click', initiateTeamsAuth);
    teamsDisconnectBtn.addEventListener('click', disconnectTeams);
    outlookAuthBtn.addEventListener('click', initiateOutlookAuth);
    outlookDisconnectBtn.addEventListener('click', disconnectOutlook);
    outlookReauthBtn.addEventListener('click', initiateOutlookAuth);
    timerTaskSaveBtn.addEventListener('click', createTimerTask);
    timerTaskRefreshBtn.addEventListener('click', loadTimerTasks);
    timerTaskList.addEventListener('click', async (event) => {
        const button = event.target.closest('.timer-task-action-btn');
        if (!button) return;
        const taskId = button.dataset.taskId;
        const action = button.dataset.action;
        if (!taskId || !action) return;

        if (action === 'run') {
            await runTimerTaskNow(taskId);
            return;
        }
        if (action === 'toggle') {
            const enabled = button.dataset.enabled === 'true';
            await updateTimerTask(taskId, { enabled: !enabled });
            return;
        }
        if (action === 'delete') {
            await deleteTimerTask(taskId);
        }
    });

    // Quick action buttons (all panels)
    document.querySelectorAll('.quick-action-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            const prompt = btn.dataset.prompt;
            if (!prompt) return;
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
    const panels = { gmail: gmailPanel, calendar: calendarPanel, gchat: gchatPanel, drive: drivePanel, sheets: sheetsPanel, docs: docsPanel, github: githubPanel, outlook: outlookPanel, teams: teamsPanel, timer: timerPanel };
    const navItems = { gmail: gmailNavItem, calendar: calendarNavItem, gchat: gchatNavItem, drive: driveNavItem, sheets: sheetsNavItem, docs: docsNavItem, github: githubNavItem, outlook: outlookNavItem, teams: teamsNavItem, timer: timerNavItem };
    if (panels[service]) {
        panels[service].classList.add('active');
        navItems[service].classList.add('active');
    }
}

function closeAllPanels() {
    [gmailPanel, calendarPanel, gchatPanel, drivePanel, sheetsPanel, docsPanel, githubPanel, outlookPanel, teamsPanel, timerPanel].forEach(p => p.classList.remove('active'));
    [gmailNavItem, calendarNavItem, gchatNavItem, driveNavItem, sheetsNavItem, docsNavItem, githubNavItem, outlookNavItem, teamsNavItem, timerNavItem].forEach(n => n.classList.remove('active'));
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
            drive: { label: 'Google Drive', dot: 'drive', categories: DRIVE_CATEGORIES },
            sheets: { label: 'Google Sheets', dot: 'sheets', categories: SHEETS_CATEGORIES },
            github: { label: 'GitHub', dot: 'github', categories: GITHUB_CATEGORIES },
            outlook: { label: 'Outlook', dot: 'outlook', categories: OUTLOOK_CATEGORIES },
            docs: { label: 'Google Docs', dot: 'docs', categories: DOCS_CATEGORIES },
            teams: { label: 'Microsoft Teams', dot: 'teams', categories: TEAMS_CATEGORIES }
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
    checkDriveStatus();
    checkSheetsStatus();
    checkDocsStatus();
    checkGitHubStatus();
    checkOutlookStatus();
    checkTeamsStatus();
    checkTimerStatus();

    setInterval(() => {
        checkGmailStatus();
        checkCalendarStatus();
        checkGchatStatus();
        checkDriveStatus();
        checkSheetsStatus();
        checkDocsStatus();
        checkGitHubStatus();
        checkOutlookStatus();
        checkTeamsStatus();
        checkTimerStatus();
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

// Google Drive status
async function checkDriveStatus() {
    try {
        const response = await fetch('/api/drive/status');
        const data = await response.json();
        updateDriveStatus(data);
    } catch (error) {
        updateDriveStatus({ authenticated: false });
    }
}

function updateDriveStatus(data) {
    const statusDot = driveStatus.querySelector('.status-dot');
    isDriveConnected = data.authenticated;

    if (data.authenticated) {
        statusDot.className = 'status-dot connected';
        driveAuthSection.style.display = 'none';
        driveConnectedSection.style.display = 'block';
        driveBadge.style.display = 'inline-flex';
    } else {
        statusDot.className = 'status-dot disconnected';
        driveAuthSection.style.display = 'block';
        driveConnectedSection.style.display = 'none';
        driveBadge.style.display = 'none';
    }
}

// Google Sheets status
async function checkSheetsStatus() {
    try {
        const response = await fetch('/api/sheets/status');
        const data = await response.json();
        updateSheetsStatus(data);
    } catch (error) {
        updateSheetsStatus({ authenticated: false });
    }
}

function updateSheetsStatus(data) {
    const statusDot = sheetsStatus.querySelector('.status-dot');
    isSheetsConnected = data.authenticated;

    if (data.authenticated) {
        statusDot.className = 'status-dot connected';
        sheetsAuthSection.style.display = 'none';
        sheetsConnectedSection.style.display = 'block';
        sheetsBadge.style.display = 'inline-flex';
    } else {
        statusDot.className = 'status-dot disconnected';
        sheetsAuthSection.style.display = 'block';
        sheetsConnectedSection.style.display = 'none';
        sheetsBadge.style.display = 'none';
    }
}

// Google OAuth (shared across Google integrations)
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
                    checkDriveStatus();
                    checkSheetsStatus();
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
                    checkDriveStatus();
                    checkSheetsStatus();
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
                    checkDriveStatus();
                    checkSheetsStatus();
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

async function initiateDriveAuth() {
    try {
        driveAuthBtn.disabled = true;
        driveAuthBtn.innerHTML = `<svg class="spinner" width="20" height="20" viewBox="0 0 24 24"><circle cx="12" cy="12" r="10" stroke="currentColor" stroke-width="3" fill="none" stroke-dasharray="30 70" /></svg> Connecting...`;

        const response = await fetch('/api/drive/connect');
        const contentType = (response.headers.get('content-type') || '').toLowerCase();
        let data = {};
        if (contentType.includes('application/json')) {
            data = await response.json();
        } else {
            const text = await response.text();
            if (response.status === 404) {
                throw new Error('Google Drive auth route not found on backend. Restart the server so latest routes are loaded.');
            }
            throw new Error(`Unexpected non-JSON response from backend (${response.status}): ${text.slice(0, 120)}`);
        }

        if (!response.ok) {
            throw new Error(data.error || `Google Drive auth failed with status ${response.status}`);
        }

        if (data.authUrl) {
            const popup = window.open(data.authUrl, 'Google Drive Auth', 'width=600,height=700,left=200,top=100');
            const checkClosed = setInterval(() => {
                if (popup.closed) {
                    clearInterval(checkClosed);
                    checkGmailStatus();
                    checkCalendarStatus();
                    checkGchatStatus();
                    checkDriveStatus();
                    checkSheetsStatus();
                    resetDriveAuthButton();
                }
            }, 500);
        } else if (data.setupRequired) {
            alert('Please set up Google Cloud credentials first.');
            resetDriveAuthButton();
        } else {
            alert(data.error || 'Failed to initiate Google Drive authentication');
            resetDriveAuthButton();
        }
    } catch (error) {
        console.error('Google Drive auth error:', error);
        alert(error.message || 'Failed to initiate Google Drive authentication');
        resetDriveAuthButton();
    }
}

function resetDriveAuthButton() {
    driveAuthBtn.disabled = false;
    driveAuthBtn.innerHTML = `<svg viewBox="0 0 24 24" width="20" height="20"><path fill="#fff" d="M22.56 12.25c0-.78-.07-1.53-.2-2.25H12v4.26h5.92c-.26 1.37-1.04 2.53-2.21 3.31v2.77h3.57c2.08-1.92 3.28-4.74 3.28-8.09z"/><path fill="#fff" d="M12 23c2.97 0 5.46-.98 7.28-2.66l-3.57-2.77c-.98.66-2.23 1.06-3.71 1.06-2.86 0-5.29-1.93-6.16-4.53H2.18v2.84C3.99 20.53 7.7 23 12 23z"/><path fill="#fff" d="M5.84 14.09c-.22-.66-.35-1.36-.35-2.09s.13-1.43.35-2.09V7.07H2.18C1.43 8.55 1 10.22 1 12s.43 3.45 1.18 4.93l2.85-2.22.81-.62z"/><path fill="#fff" d="M12 5.38c1.62 0 3.06.56 4.21 1.64l3.15-3.15C17.45 2.09 14.97 1 12 1 7.7 1 3.99 3.47 2.18 7.07l3.66 2.84c.87-2.6 3.3-4.53 6.16-4.53z"/></svg> Sign in with Google`;
}

async function initiateSheetsAuth() {
    try {
        sheetsAuthBtn.disabled = true;
        sheetsAuthBtn.innerHTML = `<svg class="spinner" width="20" height="20" viewBox="0 0 24 24"><circle cx="12" cy="12" r="10" stroke="currentColor" stroke-width="3" fill="none" stroke-dasharray="30 70" /></svg> Connecting...`;

        const response = await fetch('/api/sheets/connect');
        const contentType = (response.headers.get('content-type') || '').toLowerCase();
        let data = {};
        if (contentType.includes('application/json')) {
            data = await response.json();
        } else {
            const text = await response.text();
            if (response.status === 404) {
                throw new Error('Google Sheets auth route not found on backend. Restart the server so latest routes are loaded.');
            }
            throw new Error(`Unexpected non-JSON response from backend (${response.status}): ${text.slice(0, 120)}`);
        }

        if (!response.ok) {
            throw new Error(data.error || `Google Sheets auth failed with status ${response.status}`);
        }

        if (data.authUrl) {
            const popup = window.open(data.authUrl, 'Google Sheets Auth', 'width=600,height=700,left=200,top=100');
            const checkClosed = setInterval(() => {
                if (popup.closed) {
                    clearInterval(checkClosed);
                    checkGmailStatus();
                    checkCalendarStatus();
                    checkGchatStatus();
                    checkDriveStatus();
                    checkSheetsStatus();
                    resetSheetsAuthButton();
                }
            }, 500);
        } else if (data.setupRequired) {
            alert('Please set up Google Cloud credentials first.');
            resetSheetsAuthButton();
        } else {
            alert(data.error || 'Failed to initiate Google Sheets authentication');
            resetSheetsAuthButton();
        }
    } catch (error) {
        console.error('Google Sheets auth error:', error);
        alert(error.message || 'Failed to initiate Google Sheets authentication');
        resetSheetsAuthButton();
    }
}

function resetSheetsAuthButton() {
    sheetsAuthBtn.disabled = false;
    sheetsAuthBtn.innerHTML = `<svg viewBox="0 0 24 24" width="20" height="20"><path fill="#fff" d="M22.56 12.25c0-.78-.07-1.53-.2-2.25H12v4.26h5.92c-.26 1.37-1.04 2.53-2.21 3.31v2.77h3.57c2.08-1.92 3.28-4.74 3.28-8.09z"/><path fill="#fff" d="M12 23c2.97 0 5.46-.98 7.28-2.66l-3.57-2.77c-.98.66-2.23 1.06-3.71 1.06-2.86 0-5.29-1.93-6.16-4.53H2.18v2.84C3.99 20.53 7.7 23 12 23z"/><path fill="#fff" d="M5.84 14.09c-.22-.66-.35-1.36-.35-2.09s.13-1.43.35-2.09V7.07H2.18C1.43 8.55 1 10.22 1 12s.43 3.45 1.18 4.93l2.85-2.22.81-.62z"/><path fill="#fff" d="M12 5.38c1.62 0 3.06.56 4.21 1.64l3.15-3.15C17.45 2.09 14.97 1 12 1 7.7 1 3.99 3.47 2.18 7.07l3.66 2.84c.87-2.6 3.3-4.53 6.16-4.53z"/></svg> Sign in with Google`;
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

// Outlook status
async function checkOutlookStatus() {
    try {
        const response = await fetch('/api/outlook/status');
        const data = await response.json();
        updateOutlookStatus(data);
    } catch (error) {
        updateOutlookStatus({ authenticated: false });
    }
}

function updateOutlookStatus(data) {
    const statusDot = outlookStatus.querySelector('.status-dot');
    isOutlookConnected = data.authenticated;
    const emailText = data.email ? `Connected as ${data.email}. 18 tools ready.` : '18 tools ready.';
    document.getElementById('outlookUserEmailDisplay').textContent = emailText;

    if (!data.oauthConfigured && outlookAuthNote) {
        outlookAuthNote.textContent = 'Outlook OAuth is not configured. Add OUTLOOK_CLIENT_ID and OUTLOOK_CLIENT_SECRET to .env.';
    } else if (outlookAuthNote) {
        outlookAuthNote.textContent = 'Uses secure Microsoft OAuth. You can reauthenticate later to switch accounts.';
    }

    if (data.authenticated) {
        statusDot.className = 'status-dot connected';
        outlookAuthSection.style.display = 'none';
        outlookConnectedSection.style.display = 'block';
        outlookBadge.style.display = 'inline-flex';
    } else {
        statusDot.className = 'status-dot disconnected';
        outlookAuthSection.style.display = 'block';
        outlookConnectedSection.style.display = 'none';
        outlookBadge.style.display = 'none';
    }
}

async function initiateOutlookAuth() {
    try {
        outlookAuthBtn.disabled = true;
        outlookAuthBtn.innerHTML = '<svg class="spinner" width="20" height="20" viewBox="0 0 24 24"><circle cx="12" cy="12" r="10" stroke="currentColor" stroke-width="3" fill="none" stroke-dasharray="30 70" /></svg> Connecting...';

        const response = await fetch('/api/outlook/auth');
        const contentType = (response.headers.get('content-type') || '').toLowerCase();
        let data = {};
        if (contentType.includes('application/json')) {
            data = await response.json();
        } else {
            const text = await response.text();
            throw new Error(`Unexpected response (${response.status}): ${text.slice(0, 120)}`);
        }

        if (!response.ok) {
            throw new Error(data.error || `Outlook auth failed with status ${response.status}`);
        }

        if (data.authUrl) {
            const popup = window.open(data.authUrl, 'Outlook Authentication', 'width=600,height=700,left=200,top=100');
            const checkClosed = setInterval(() => {
                if (popup.closed) {
                    clearInterval(checkClosed);
                    checkOutlookStatus();
                    resetOutlookAuthButton();
                }
            }, 500);
        } else if (data.setupRequired) {
            alert('Please configure OUTLOOK_CLIENT_ID and OUTLOOK_CLIENT_SECRET in .env first.');
            resetOutlookAuthButton();
        } else {
            alert(data.error || 'Failed to start Outlook authentication');
            resetOutlookAuthButton();
        }
    } catch (error) {
        console.error('Outlook OAuth error:', error);
        alert(error.message || 'Failed to initiate Outlook authentication');
        resetOutlookAuthButton();
    }
}

function resetOutlookAuthButton() {
    outlookAuthBtn.disabled = false;
    outlookAuthBtn.innerHTML = '<svg viewBox="0 0 21 21" width="20" height="20"><rect x="1" y="1" width="9" height="9" fill="#f25022"/><rect x="11" y="1" width="9" height="9" fill="#7fba00"/><rect x="1" y="11" width="9" height="9" fill="#00a4ef"/><rect x="11" y="11" width="9" height="9" fill="#ffb900"/></svg> Sign in with Microsoft';
}

async function disconnectOutlook() {
    try {
        await fetch('/api/outlook/disconnect', { method: 'POST' });
        checkOutlookStatus();
    } catch (error) {
        console.error('Outlook disconnect error:', error);
    }
}

// Google Docs status
async function checkDocsStatus() {
    try {
        const response = await fetch('/api/docs/status');
        const data = await response.json();
        updateDocsStatus(data);
    } catch (error) {
        updateDocsStatus({ authenticated: false });
    }
}

function updateDocsStatus(data) {
    const statusDot = docsStatus.querySelector('.status-dot');
    isDocsConnected = data.authenticated;
    if (data.authenticated) {
        statusDot.className = 'status-dot connected';
        docsAuthSection.style.display = 'none';
        docsConnectedSection.style.display = 'block';
        docsBadge.style.display = 'inline-flex';
    } else {
        statusDot.className = 'status-dot disconnected';
        docsAuthSection.style.display = 'block';
        docsConnectedSection.style.display = 'none';
        docsBadge.style.display = 'none';
    }
}

// Microsoft Teams status
async function checkTeamsStatus() {
    try {
        const response = await fetch('/api/teams/status');
        const data = await response.json();
        updateTeamsStatus(data);
    } catch (error) {
        updateTeamsStatus({ authenticated: false, oauthConfigured: false });
    }
}

function updateTeamsStatus(data) {
    const statusDot = teamsStatus.querySelector('.status-dot');
    isTeamsConnected = data.authenticated;
    if (data.authenticated) {
        statusDot.className = 'status-dot connected';
        teamsAuthSection.style.display = 'none';
        teamsConnectedSection.style.display = 'block';
        teamsSetupSection.style.display = 'none';
        teamsBadge.style.display = 'inline-flex';
        if (data.email) teamsUserInfo.textContent = `Connected as ${data.email}. 10 team tools ready!`;
    } else if (!data.oauthConfigured) {
        statusDot.className = 'status-dot disconnected';
        teamsAuthSection.style.display = 'none';
        teamsConnectedSection.style.display = 'none';
        teamsSetupSection.style.display = 'block';
        teamsBadge.style.display = 'none';
    } else {
        statusDot.className = 'status-dot disconnected';
        teamsAuthSection.style.display = 'block';
        teamsConnectedSection.style.display = 'none';
        teamsSetupSection.style.display = 'none';
        teamsBadge.style.display = 'none';
    }
}

async function initiateTeamsAuth() {
    try {
        const response = await fetch('/api/teams/auth');
        const data = await response.json();
        if (data.authUrl) {
            const popup = window.open(data.authUrl, 'TeamsAuth', 'width=600,height=700');
            const pollInterval = setInterval(() => {
                if (popup.closed) {
                    clearInterval(pollInterval);
                    setTimeout(() => {
                        checkTeamsStatus();
                        checkOutlookStatus();
                    }, 1000);
                }
            }, 500);
        } else {
            alert(data.error || 'Failed to get Teams auth URL');
        }
    } catch (error) {
        alert('Failed to initiate Teams authentication: ' + error.message);
    }
}

async function disconnectTeams() {
    try {
        await fetch('/api/outlook/disconnect', { method: 'POST' });
        checkTeamsStatus();
        checkOutlookStatus();
    } catch (error) {
        console.error('Teams disconnect error:', error);
    }
}

async function checkTimerStatus() {
    try {
        const response = await fetch('/api/timer-tasks/status');
        const data = await response.json();
        updateTimerStatus(data);
    } catch (error) {
        updateTimerStatus({ connected: false, taskCount: 0, enabledCount: 0, runningCount: 0 });
    }
}

function updateTimerStatus(data) {
    const statusDot = timerStatus.querySelector('.status-dot');
    const enabledCount = Number(data.enabledCount || 0);
    const taskCount = Number(data.taskCount || 0);
    isTimerConnected = !!data.connected;

    statusDot.className = `status-dot ${enabledCount > 0 ? 'connected' : 'disconnected'}`;
    timerBadge.style.display = enabledCount > 0 ? 'inline-flex' : 'none';

    if (timerTaskStatusText) {
        if (taskCount === 0) {
            timerTaskStatusText.textContent = 'No tasks configured yet.';
        } else {
            timerTaskStatusText.textContent = `${taskCount} task(s), ${enabledCount} enabled, ${Number(data.runningCount || 0)} running now.`;
        }
    }
}

async function loadTimerTasks() {
    try {
        const response = await fetch('/api/timer-tasks');
        const data = await response.json();
        if (!response.ok) {
            throw new Error(data.error || `Failed to load timer tasks (${response.status})`);
        }
        renderTimerTaskList(Array.isArray(data.tasks) ? data.tasks : []);
        await checkTimerStatus();
    } catch (error) {
        console.error('Timer task load error:', error);
        timerTaskList.innerHTML = `<div class="timer-task-meta" style="color: var(--error);">Failed to load tasks: ${escapeHtml(error.message)}</div>`;
    }
}

function renderTimerTaskList(tasks) {
    if (!tasks || tasks.length === 0) {
        timerTaskList.innerHTML = '<div class="timer-task-meta">No timer tasks yet. Create one above.</div>';
        return;
    }
    timerTaskList.innerHTML = tasks.map(task => `
        <div class="timer-task-card">
            <div class="timer-task-header">
                <span class="timer-task-name">${escapeHtml(task.name || 'Untitled task')}</span>
                <span class="timer-task-time">${escapeHtml(task.time || '')}</span>
            </div>
            <div class="timer-task-meta">Status: ${escapeHtml(task.lastStatus || 'never')} ${task.lastRunAt ? `- Last run ${escapeHtml(formatDate(task.lastRunAt))}` : ''} ${task.running ? '- Running...' : ''}</div>
            ${task.lastError ? `<div class="timer-task-meta" style="color: var(--error);">Error: ${escapeHtml(task.lastError)}</div>` : ''}
            <div class="timer-task-instruction">${escapeHtml(task.instruction || '')}</div>
            <div class="timer-task-actions">
                <button class="timer-task-action-btn" data-action="run" data-task-id="${escapeHtml(task.id)}">Run Now</button>
                <button class="timer-task-action-btn" data-action="toggle" data-enabled="${task.enabled ? 'true' : 'false'}" data-task-id="${escapeHtml(task.id)}">${task.enabled ? 'Disable' : 'Enable'}</button>
                <button class="timer-task-action-btn" data-action="delete" data-task-id="${escapeHtml(task.id)}">Delete</button>
            </div>
        </div>
    `).join('');
}

async function createTimerTask() {
    const instruction = (timerTaskInstructionInput.value || '').trim();
    const time = (timerTaskTimeInput.value || '').trim();
    const name = (timerTaskNameInput.value || '').trim();
    const enabled = !!timerTaskEnabledInput.checked;

    if (!instruction) {
        alert('Please enter task instructions.');
        return;
    }
    if (!time) {
        alert('Please choose a daily time.');
        return;
    }

    timerTaskSaveBtn.disabled = true;
    const originalLabel = timerTaskSaveBtn.innerHTML;
    timerTaskSaveBtn.innerHTML = '&#9201; Saving...';
    try {
        const response = await fetch('/api/timer-tasks', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ name, instruction, time, enabled })
        });
        const data = await response.json();
        if (!response.ok) throw new Error(data.error || `Failed to create task (${response.status})`);

        timerTaskNameInput.value = '';
        timerTaskInstructionInput.value = '';
        timerTaskEnabledInput.checked = true;
        await loadTimerTasks();
    } catch (error) {
        alert(error.message || 'Failed to create timer task');
    } finally {
        timerTaskSaveBtn.disabled = false;
        timerTaskSaveBtn.innerHTML = originalLabel;
    }
}

async function updateTimerTask(taskId, patch) {
    try {
        const response = await fetch(`/api/timer-tasks/${encodeURIComponent(taskId)}`, {
            method: 'PATCH',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(patch)
        });
        const data = await response.json();
        if (!response.ok) throw new Error(data.error || `Failed to update task (${response.status})`);
        await loadTimerTasks();
    } catch (error) {
        alert(error.message || 'Failed to update timer task');
    }
}

async function deleteTimerTask(taskId) {
    if (!confirm('Delete this timer task?')) return;
    try {
        const response = await fetch(`/api/timer-tasks/${encodeURIComponent(taskId)}`, { method: 'DELETE' });
        const data = await response.json();
        if (!response.ok) throw new Error(data.error || `Failed to delete task (${response.status})`);
        await loadTimerTasks();
    } catch (error) {
        alert(error.message || 'Failed to delete timer task');
    }
}

async function runTimerTaskNow(taskId) {
    try {
        const response = await fetch(`/api/timer-tasks/${encodeURIComponent(taskId)}/run`, { method: 'POST' });
        const data = await response.json();
        if (!response.ok) throw new Error(data.error || `Failed to run task (${response.status})`);
        await loadTimerTasks();
        if (data.skipped) {
            alert('Task is already running.');
        }
    } catch (error) {
        alert(error.message || 'Failed to run timer task');
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
            addMessage('assistant', `<p style="color: #ef4444;">Error: ${escapeHtml(data.error)}</p>`, { allowHtml: true });
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

            // Suppress text response for list-type tools (only show cards)
            const LIST_TOOLS = ['list_emails', 'search_emails', 'list_events', 'list_repos', 'list_issues', 'list_prs', 'list_drive_files', 'list_spreadsheets', 'list_chat_spaces'];
            const isListToolOnly = data.toolResults &&
                data.toolResults.length === 1 &&
                LIST_TOOLS.includes(data.toolResults[0].tool) &&
                data.toolResults[0].result;

            if (!isListToolOnly) {
                responseHtml += formatResponse(data.response);
            }

            if (data.toolResults && data.toolResults.length > 0) {
                responseHtml += formatToolResults(data.toolResults);
            }

            addMessage('assistant', responseHtml, { allowHtml: true });

            chatHistory.push({ role: 'user', content: message });
            chatHistory.push({ role: 'assistant', content: data.response });

            if (chatHistory.length > 30) {
                chatHistory = chatHistory.slice(-30);
            }
        }
    } catch (error) {
        removeTypingIndicator(typingId);
        addMessage('assistant', `<p style="color: #ef4444;">Failed to send message. Please try again.</p>`, { allowHtml: true });
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
                content = result.result.results.map(email => {
                    const sender = email.from || 'Unknown';
                    const nameMatch = sender.match(/^([^<]+)/);
                    const displayName = nameMatch ? nameMatch[1].trim().replace(/["']/g, '') : sender;
                    const initials = (displayName[0] || '?').toUpperCase();

                    return `
                    <div class="email-card email-clickable" data-message-id="${escapeHtml(email.id || '')}" onclick="openEmail('${escapeHtml(email.id || '')}', this)">
                        <div class="email-card-content-wrapper">
                            <div class="email-avatar">${escapeHtml(initials)}</div>
                            <div class="email-main-info">
                                <div class="email-sender">${escapeHtml(displayName)}</div>
                                <div class="email-snippet-wrapper">
                                    <span class="email-subject">${escapeHtml(email.subject || '(no subject)')}</span>
                                    <span class="email-snippet">- ${escapeHtml(email.snippet ? email.snippet.slice(0, 90) : '')}...</span>
                                </div>
                                <div class="email-date">${formatDate(email.date)}</div>
                            </div>
                        </div>
                    </div>
                `}).join('');
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
                // Drive files
            } else if (result.result.files && Array.isArray(result.result.files)) {
                content = result.result.files.map(f => `
                    <div class="email-card">
                        <div class="email-card-header">
                            <span class="email-card-subject">${escapeHtml(f.name || f.id)}</span>
                            <span class="email-card-date">${escapeHtml((f.mimeType || '').replace('application/vnd.google-apps.', ''))}</span>
                        </div>
                        <div class="email-card-from"><code>${escapeHtml(f.id || '')}</code></div>
                        ${f.webViewLink ? `<div class="email-card-snippet"><a href="${escapeHtml(f.webViewLink)}" target="_blank" rel="noopener noreferrer">Open in Drive</a></div>` : ''}
                    </div>
                `).join('');
                // Spreadsheet list
            } else if (result.result.spreadsheets && Array.isArray(result.result.spreadsheets)) {
                content = result.result.spreadsheets.map(s => `
                    <div class="email-card">
                        <div class="email-card-header">
                            <span class="email-card-subject">${escapeHtml(s.title || s.spreadsheetId)}</span>
                            <span class="email-card-date">${formatDate(s.modifiedTime)}</span>
                        </div>
                        <div class="email-card-from"><code>${escapeHtml(s.spreadsheetId || '')}</code></div>
                        ${s.webViewLink ? `<div class="email-card-snippet"><a href="${escapeHtml(s.webViewLink)}" target="_blank" rel="noopener noreferrer">Open Spreadsheet</a></div>` : ''}
                    </div>
                `).join('');
                // Sheet tabs
            } else if (result.result.tabs && Array.isArray(result.result.tabs)) {
                content = `<div class="labels-list">${result.result.tabs.map(t => `<span class="label-chip">${escapeHtml(t.title || String(t.sheetId))}</span>`).join('')}</div>`;
                // Sheet values
            } else if (result.result.values && Array.isArray(result.result.values)) {
                const preview = result.result.values.slice(0, 20);
                content = `<pre style="font-size:0.8rem;max-height:180px;overflow:auto">${escapeHtml(JSON.stringify(preview, null, 2))}</pre>`;
                // GitHub repos
            } else if (result.result.repos && Array.isArray(result.result.repos)) {
                content = result.result.repos.map(r => `
                    <div class="email-card vertical">
                        <div class="email-card-header">
                            <span class="email-card-subject">${escapeHtml(r.full_name || r.name)}</span>
                            <span class="email-card-date">${r.language || ''}</span>
                        </div>
                        ${r.description ? `<div class="email-card-snippet">${escapeHtml(r.description.slice(0, 100))}</div>` : ''}
                        <div class="email-card-from" style="margin-top:0.25rem">&#11088; ${r.stars || 0} &middot; &#128204; ${r.forks || 0}${r.private ? ' &middot; Private' : ''}</div>
                        ${r.url ? `<div class="email-card-snippet" style="margin-top:0.5rem"><a href="${escapeHtml(r.url)}" target="_blank" rel="noopener noreferrer" style="color:var(--accent-primary);text-decoration:none;font-weight:600">Go to Repo &rarr;</a></div>` : ''}
                    </div>
                `).join('');
                // GitHub issues
            } else if (result.result.issues && Array.isArray(result.result.issues)) {
                content = result.result.issues.map(i => `
                    <div class="email-card vertical">
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
                    <div class="email-card vertical">
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
                    <div class="email-card vertical">
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
                    <div class="email-card vertical">
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
                    <div class="email-card vertical">
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

// Open email when card is clicked
async function openEmail(messageId, cardElement) {
    if (!messageId || !cardElement) return;

    // Check if already expanded
    const existingDetails = cardElement.querySelector('.email-details');
    if (existingDetails) {
        existingDetails.remove();
        cardElement.classList.remove('expanded');
        return;
    }

    // Add loading placeholder
    const detailsDiv = document.createElement('div');
    detailsDiv.className = 'email-details loading';
    detailsDiv.innerHTML = '<div class="spinner" style="width:20px;height:20px;border-width:2px;margin-right:10px"></div> Loading content...';
    cardElement.appendChild(detailsDiv);
    cardElement.classList.add('expanded');

    try {
        const response = await fetch(`/api/gmail/message/${messageId}`);
        const data = await response.json();

        if (data.error) throw new Error(data.error);

        // Render content
        detailsDiv.classList.remove('loading');

        const sender = data.from || 'Unknown';
        const initials = (sender.replace(/["']/g, '').trim()[0] || '?').toUpperCase();

        let attachmentsHtml = '';
        if (data.attachments && data.attachments.length > 0) {
            attachmentsHtml = `
                <div class="email-attachments" style="margin-top: 1rem; padding-top: 1rem; border-top: 1px dashed var(--border-color);">
                    <strong>Attachments:</strong>
                    <div class="attachment-list" style="display:flex; gap:0.5rem; flex-wrap:wrap; margin-top:0.5rem">
                        ${data.attachments.map(a => `
                            <div class="attachment-chip" style="background:var(--bg-tertiary); padding:0.25rem 0.5rem; border-radius:4px; font-size:0.85rem; border:1px solid var(--border-color)">
                                &#128206; ${escapeHtml(a.filename)} (${formatSize(a.size)})
                            </div>
                        `).join('')}
                    </div>
                </div>
            `;
        }

        detailsDiv.innerHTML = `
            <div class="email-details-content">
                <div class="email-details-header">
                    <div style="flex:1">
                        <div class="email-details-subject">${escapeHtml(data.subject || '(no subject)')}</div>
                        <div class="email-details-meta">
                            <div class="email-avatar">${escapeHtml(initials)}</div>
                            <div class="email-details-sender-info">
                                <span class="email-details-from">${escapeHtml(sender)}</span>
                                <span class="email-details-to">to ${escapeHtml(data.to || 'me')}</span>
                            </div>
                            <div class="email-date" style="margin-left:auto">${new Date(data.date).toLocaleString()}</div>
                        </div>
                    </div>
                </div>
                
                <hr style="margin: 1rem 0; border: 0; border-top: 1px solid var(--border-color); opacity:0.5">
                
                <div class="email-details-body">${data.bodyHtml || escapeHtml(data.body || '(No content)')}</div>
                
                ${attachmentsHtml}

                <div class="email-actions" style="margin-top: 1.5rem; display: flex; gap: 0.5rem;">
                    <button class="action-btn" style="padding:0.5rem 1rem; background:var(--accent-primary); color:white; border:none; border-radius:6px; cursor:pointer" onclick="event.stopPropagation(); messageInput.value='Reply to email ${messageId} saying...'; messageInput.focus();">Reply</button>
                    <button class="action-btn" style="padding:0.5rem 1rem; background:var(--bg-tertiary); color:var(--text-primary); border:1px solid var(--border-color); border-radius:6px; cursor:pointer" onclick="event.stopPropagation(); messageInput.value='Forward email ${messageId} to...'; messageInput.focus();">Forward</button>
                </div>
            </div>
        `;
    } catch (error) {
        detailsDiv.innerHTML = `<div style="color:var(--error); padding:1rem">Error loading email: ${escapeHtml(error.message)}</div>`;
    }
}

// Add message to chat
function addMessage(role, content, { allowHtml = false } = {}) {
    const messageDiv = document.createElement('div');
    messageDiv.className = `message ${role}`;

    const avatar = role === 'user'
        ? '<svg viewBox="0 0 24 24" width="20" height="20"><path fill="#fff" d="M12 12c2.21 0 4-1.79 4-4s-1.79-4-4-4-4 1.79-4 4 1.79 4 4 4zm0 2c-2.67 0-8 1.34-8 4v2h16v-2c0-2.66-5.33-4-8-4z"/></svg>'
        : '<svg viewBox="0 0 24 24" width="20" height="20"><circle cx="12" cy="12" r="10" fill="none" stroke="#6366f1" stroke-width="2"/><path fill="none" stroke="#6366f1" stroke-width="2" d="M8 12l3 3 5-6"/></svg>';

    const bubbleContent = allowHtml
        ? String(content || '')
        : escapeHtml(String(content || '')).replace(/\n/g, '<br>');

    messageDiv.innerHTML = `
        <div class="message-avatar">${avatar}</div>
        <div class="message-content">
            <div class="message-bubble">${bubbleContent}</div>
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
