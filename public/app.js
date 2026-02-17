// DOM Elements
const chatMessages = document.getElementById('chatMessages');
const messageInput = document.getElementById('messageInput');
const micBtn = document.getElementById('micBtn');
const sendBtn = document.getElementById('sendBtn');
const attachBtn = document.getElementById('attachBtn');
const fileInput = document.getElementById('fileInput');
const attachedFilesPreview = document.getElementById('attachedFilesPreview');
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
const gcsNavItem = document.getElementById('gcsNavItem');
const gcsPanel = document.getElementById('gcsPanel');
const gcsSetupSection = document.getElementById('gcsSetupSection');
const gcsConnectedSection = document.getElementById('gcsConnectedSection');
const gcsStatus = document.getElementById('gcsStatus');
const gcsBadge = document.getElementById('gcsBadge');
const gcsProjectInfo = document.getElementById('gcsProjectInfo');
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
const CHAT_HISTORY_MAX_MESSAGES = 24;
const CHAT_HISTORY_MAX_MESSAGE_CHARS = 3000;
let isGmailConnected = false;
let isCalendarConnected = false;
let isGchatConnected = false;
let isDriveConnected = false;
let isSheetsConnected = false;
let isGithubConnected = false;
let isOutlookConnected = false;
let isDocsConnected = false;
let isTeamsConnected = false;
let isGcsConnected = false;
let isTimerConnected = false;
let isRecording = false;
let recognition = null;
let activeFilter = 'all';

function truncateHistoryText(text, maxChars = CHAT_HISTORY_MAX_MESSAGE_CHARS) {
    const value = String(text ?? '');
    if (value.length <= maxChars) return value;
    if (maxChars <= 20) return value.slice(0, Math.max(0, maxChars));
    return `${value.slice(0, maxChars - 20)}\n...[truncated]`;
}

function pushChatHistoryEntry(role, content) {
    const clean = truncateHistoryText(String(content ?? '').trim(), CHAT_HISTORY_MAX_MESSAGE_CHARS);
    if (!clean) return;
    chatHistory.push({ role, content: clean });
    if (chatHistory.length > CHAT_HISTORY_MAX_MESSAGES) {
        chatHistory = chatHistory.slice(-CHAT_HISTORY_MAX_MESSAGES);
    }
}

function triggerBrowserDownload(url, filename) {
    const href = String(url || '').trim();
    if (!href) return;
    try {
        // Security: only allow http/https/blob URLs
        const parsed = new URL(href, window.location.origin);
        if (!['http:', 'https:', 'blob:'].includes(parsed.protocol)) return;
        const anchor = document.createElement('a');
        anchor.href = href;
        if (filename) anchor.setAttribute('download', String(filename));
        anchor.rel = 'noopener noreferrer';
        anchor.style.display = 'none';
        document.body.appendChild(anchor);
        anchor.click();
        setTimeout(() => anchor.remove(), 0);
    } catch (error) {
        console.warn('Download trigger failed:', error);
    }
}

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
    download_drive_file: '&#128229;', extract_drive_file_text: '&#128214;', append_drive_document_text: '&#9998;', convert_file_to_google_doc: '&#128260;', convert_file_to_google_sheet: '&#128202;',
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
    list_commits: '&#128221;', revert_commit: '&#9194;', get_user_profile: '&#128100;',
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
    teams_create_chat: '&#10133;', teams_get_chat_members: '&#128101;',
    // GCS (10)
    gcs_list_buckets: '&#128230;', gcs_get_bucket: '&#128196;',
    gcs_create_bucket: '&#10133;', gcs_delete_bucket: '&#128465;',
    gcs_list_objects: '&#128193;', gcs_upload_object: '&#128228;',
    gcs_download_object: '&#128229;', gcs_delete_object: '&#128465;',
    gcs_copy_object: '&#128209;', gcs_get_object_metadata: '&#128196;'
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
    'Browse': ['list_drive_files', 'get_drive_file'],
    'Download & Extract': ['download_drive_file', 'extract_drive_file_text'],
    'Create & Convert': ['create_drive_folder', 'create_drive_file', 'append_drive_document_text', 'convert_file_to_google_doc', 'convert_file_to_google_sheet', 'update_drive_file'],
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
    'Activity': ['list_commits', 'revert_commit', 'get_user_profile', 'list_notifications', 'list_gists']
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

const GCS_CATEGORIES = {
    'Buckets': ['gcs_list_buckets', 'gcs_get_bucket', 'gcs_create_bucket', 'gcs_delete_bucket'],
    'Objects': ['gcs_list_objects', 'gcs_upload_object', 'gcs_download_object', 'gcs_delete_object', 'gcs_copy_object', 'gcs_get_object_metadata']
};

// Initialize
document.addEventListener('DOMContentLoaded', () => {
    checkAllStatuses();
    setupEventListeners();
    autoResizeTextarea();
});

// Event Listeners
function setupEventListeners() {
    setupSpeechRecognition();

    // Event delegation for dynamically created email cards (security: no inline onclick)
    chatMessages.addEventListener('click', (e) => {
        const emailCard = e.target.closest('.email-clickable[data-message-id]');
        if (emailCard) openEmail(emailCard.dataset.messageId, emailCard);
    });

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
    gcsNavItem.addEventListener('click', () => openPanel('gcs'));
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
        if (action === 'edit') {
            // Populate form with existing task data
            timerTaskNameInput.value = button.dataset.taskName || '';
            timerTaskInstructionInput.value = button.dataset.taskInstruction || '';
            timerTaskTimeInput.value = button.dataset.taskTime || '';
            timerTaskEnabledInput.checked = button.dataset.taskEnabled === 'true';

            // Change save button to update mode
            timerTaskSaveBtn.textContent = 'Update Task';
            timerTaskSaveBtn.dataset.editingTaskId = taskId;

            // Scroll to form
            timerTaskNameInput.scrollIntoView({ behavior: 'smooth', block: 'center' });
            timerTaskInstructionInput.focus();
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
    const panels = { gmail: gmailPanel, calendar: calendarPanel, gchat: gchatPanel, drive: drivePanel, sheets: sheetsPanel, docs: docsPanel, github: githubPanel, outlook: outlookPanel, teams: teamsPanel, gcs: gcsPanel, timer: timerPanel };
    const navItems = { gmail: gmailNavItem, calendar: calendarNavItem, gchat: gchatNavItem, drive: driveNavItem, sheets: sheetsNavItem, docs: docsNavItem, github: githubNavItem, outlook: outlookNavItem, teams: teamsNavItem, gcs: gcsNavItem, timer: timerNavItem };
    if (panels[service]) {
        panels[service].classList.add('active');
        navItems[service].classList.add('active');
    }
}

function closeAllPanels() {
    [gmailPanel, calendarPanel, gchatPanel, drivePanel, sheetsPanel, docsPanel, githubPanel, outlookPanel, teamsPanel, gcsPanel, timerPanel].forEach(p => p.classList.remove('active'));
    [gmailNavItem, calendarNavItem, gchatNavItem, driveNavItem, sheetsNavItem, docsNavItem, githubNavItem, outlookNavItem, teamsNavItem, gcsNavItem, timerNavItem].forEach(n => n.classList.remove('active'));
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
            teams: { label: 'Microsoft Teams', dot: 'teams', categories: TEAMS_CATEGORIES },
            gcs: { label: 'GCP Cloud Storage', dot: 'gcs', categories: GCS_CATEGORIES }
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
    checkGcsStatus();
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
        checkGcsStatus();
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
        const contentType = (response.headers.get('content-type') || '').toLowerCase();
        let data = {};
        if (contentType.includes('application/json')) {
            data = await response.json();
        } else {
            const text = await response.text();
            if (response.status === 404) {
                throw new Error('Google auth route not found on backend. Restart the server so latest routes are loaded.');
            }
            throw new Error(`Unexpected non-JSON response from backend (${response.status}): ${text.slice(0, 120)}`);
        }

        if (!response.ok) {
            throw new Error(data.error || `Google auth failed with status ${response.status}`);
        }

        if (data.authUrl) {
            const popup = window.open(data.authUrl, 'Google Authentication', 'width=600,height=700,left=200,top=100');
            if (!popup) {
                alert('Popup was blocked. Please allow popups for this site and try again.');
                resetGoogleAuthButton();
                return;
            }
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
        } else {
            alert(data.error || 'Failed to initiate Google authentication');
            resetGoogleAuthButton();
        }
    } catch (error) {
        console.error('Google auth error:', error);
        alert(error.message || 'Failed to initiate authentication');
        resetGoogleAuthButton();
    }
}

async function initiateCalendarAuth() {
    try {
        calendarAuthBtn.disabled = true;
        calendarAuthBtn.innerHTML = `<svg class="spinner" width="20" height="20" viewBox="0 0 24 24"><circle cx="12" cy="12" r="10" stroke="currentColor" stroke-width="3" fill="none" stroke-dasharray="30 70" /></svg> Connecting...`;

        const response = await fetch('/api/calendar/connect');
        const contentType = (response.headers.get('content-type') || '').toLowerCase();
        let data = {};
        if (contentType.includes('application/json')) {
            data = await response.json();
        } else {
            const text = await response.text();
            if (response.status === 404) {
                throw new Error('Calendar auth route not found on backend. Restart the server so latest routes are loaded.');
            }
            throw new Error(`Unexpected non-JSON response from backend (${response.status}): ${text.slice(0, 120)}`);
        }

        if (!response.ok) {
            throw new Error(data.error || `Calendar auth failed with status ${response.status}`);
        }

        if (data.authUrl) {
            const popup = window.open(data.authUrl, 'Google Calendar Auth', 'width=600,height=700,left=200,top=100');
            if (!popup) {
                alert('Popup was blocked. Please allow popups for this site and try again.');
                resetCalendarAuthButton();
                return;
            }
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
        } else {
            alert(data.error || 'Failed to initiate calendar authentication');
            resetCalendarAuthButton();
        }
    } catch (error) {
        console.error('Calendar auth error:', error);
        alert(error.message || 'Failed to initiate calendar authentication');
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
            if (!popup) {
                alert('Popup was blocked. Please allow popups for this site and try again.');
                resetGchatAuthButton();
                return;
            }
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
            if (!popup) {
                alert('Popup was blocked. Please allow popups for this site and try again.');
                resetDriveAuthButton();
                return;
            }
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
            if (!popup) {
                alert('Popup was blocked. Please allow popups for this site and try again.');
                resetSheetsAuthButton();
                return;
            }
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
            if (!popup) {
                alert('Popup was blocked. Please allow popups for this site and try again.');
                resetGithubAuthButton();
                return;
            }
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
            if (!popup) {
                alert('Popup was blocked. Please allow popups for this site and try again.');
                resetOutlookAuthButton();
                return;
            }
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
            if (!popup) {
                alert('Popup was blocked. Please allow popups for this site and try again.');
                return;
            }
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

async function checkGcsStatus() {
    try {
        const response = await fetch('/api/gcs/status');
        const data = await response.json();
        updateGcsStatus(data);
    } catch (error) {
        updateGcsStatus({ authenticated: false });
    }
}

function updateGcsStatus(data) {
    const statusDot = gcsStatus.querySelector('.status-dot');
    isGcsConnected = data.authenticated;
    if (data.authenticated) {
        statusDot.className = 'status-dot connected';
        gcsSetupSection.style.display = 'none';
        gcsConnectedSection.style.display = 'block';
        gcsBadge.style.display = 'inline-flex';
        if (data.projectId) gcsProjectInfo.textContent = `Project: ${data.projectId}. 10 bucket/object tools ready!`;
    } else {
        statusDot.className = 'status-dot disconnected';
        gcsSetupSection.style.display = 'block';
        gcsConnectedSection.style.display = 'none';
        gcsBadge.style.display = 'none';
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
                <button class="timer-task-action-btn" data-action="edit" data-task-id="${escapeHtml(task.id)}" data-task-name="${escapeHtml(task.name || '')}" data-task-instruction="${escapeHtml(task.instruction || '')}" data-task-time="${escapeHtml(task.time || '')}" data-task-enabled="${task.enabled ? 'true' : 'false'}">Edit</button>
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
    const editingTaskId = timerTaskSaveBtn.dataset.editingTaskId;

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
        // Check if we're editing or creating
        const isEditing = !!editingTaskId;
        const url = isEditing ? `/api/timer-tasks/${encodeURIComponent(editingTaskId)}` : '/api/timer-tasks';
        const method = isEditing ? 'PATCH' : 'POST';

        const response = await fetch(url, {
            method: method,
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ name, instruction, time, enabled })
        });
        const data = await response.json();
        if (!response.ok) throw new Error(data.error || `Failed to ${isEditing ? 'update' : 'create'} task (${response.status})`);

        // Clear form and reset to create mode
        timerTaskNameInput.value = '';
        timerTaskInstructionInput.value = '';
        timerTaskEnabledInput.checked = true;
        timerTaskSaveBtn.textContent = 'Create Task';
        delete timerTaskSaveBtn.dataset.editingTaskId;

        await loadTimerTasks();
    } catch (error) {
        alert(error.message || 'Failed to save timer task');
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

// ============================================================
//  FILE UPLOAD LOGIC
// ============================================================

let attachedFiles = []; // Array of { fileId, name, size, mimeType }

// Handle attach button click
attachBtn.addEventListener('click', () => {
    fileInput.click();
});

// Handle file selection
fileInput.addEventListener('change', async (e) => {
    const files = Array.from(e.target.files || []);
    if (files.length === 0) return;

    try {
        // Upload files to server
        const formData = new FormData();
        files.forEach(file => {
            formData.append('files', file);
        });

        const response = await fetch('/api/upload', {
            method: 'POST',
            body: formData
        });

        if (!response.ok) {
            throw new Error('Failed to upload files');
        }

        const data = await response.json();

        // Add uploaded files to attachedFiles array
        attachedFiles.push(...data.files);

        // Update UI
        updateAttachedFilesUI();

        // Clear file input
        fileInput.value = '';
    } catch (error) {
        console.error('Upload error:', error);
        alert('Failed to upload files. Please try again.');
    }
});

// Update attached files UI
function updateAttachedFilesUI() {
    if (attachedFiles.length === 0) {
        attachedFilesPreview.style.display = 'none';
        attachBtn.classList.remove('has-files');
        attachedFilesPreview.innerHTML = '';
        return;
    }

    attachBtn.classList.add('has-files');
    attachedFilesPreview.style.display = 'block';

    attachedFilesPreview.innerHTML = attachedFiles.map((file, index) => `
        <div class="attached-file-item" data-index="${index}">
            <div class="attached-file-icon">
                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                    <path d="M13 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V9z"/>
                    <polyline points="13 2 13 9 20 9"/>
                </svg>
            </div>
            <div class="attached-file-info">
                <div class="attached-file-name">${escapeHtml(file.name)}</div>
                <div class="attached-file-size">${formatFileSize(file.size)}</div>
            </div>
            <button class="attached-file-remove" data-file-id="${file.fileId}" title="Remove file">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                    <line x1="18" y1="6" x2="6" y2="18"/>
                    <line x1="6" y1="6" x2="18" y2="18"/>
                </svg>
            </button>
        </div>
    `).join('');

    // Add remove button event listeners
    attachedFilesPreview.querySelectorAll('.attached-file-remove').forEach(btn => {
        btn.addEventListener('click', () => removeAttachedFile(btn.dataset.fileId));
    });
}

// Remove attached file
async function removeAttachedFile(fileId) {
    try {
        // Delete from server
        await fetch(`/api/upload/${fileId}`, {
            method: 'DELETE'
        });

        // Remove from local array
        attachedFiles = attachedFiles.filter(f => f.fileId !== fileId);

        // Update UI
        updateAttachedFilesUI();
    } catch (error) {
        console.error('Failed to remove file:', error);
    }
}

// Format file size helper
function formatFileSize(bytes) {
    if (bytes < 1024) return bytes + ' B';
    if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
    return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
}

// Send message
async function sendMessage() {
    const message = messageInput.value.trim();
    if (!message) return;

    messageInput.value = '';
    autoResizeTextarea();
    sendBtn.disabled = true;

    addMessage('user', message);

    // Create assistant message bubble immediately for streaming
    const assistantMsgDiv = document.createElement('div');
    assistantMsgDiv.className = 'message assistant';
    assistantMsgDiv.innerHTML = `
        <div class="message-avatar">
            <svg viewBox="0 0 24 24" width="20" height="20">
                <circle cx="12" cy="12" r="10" fill="none" stroke="#6366f1" stroke-width="2" />
                <path fill="none" stroke="#6366f1" stroke-width="2" d="M8 12l3 3 5-6" />
            </svg>
        </div>
        <div class="message-content">
            <div class="message-bubble">
                <div class="streaming-text"></div>
                <div class="streaming-tools"></div>
            </div>
        </div>
    `;
    chatMessages.appendChild(assistantMsgDiv);
    chatMessages.scrollTop = chatMessages.scrollHeight;

    const streamingText = assistantMsgDiv.querySelector('.streaming-text');
    const streamingTools = assistantMsgDiv.querySelector('.streaming-tools');
    let accumulatedText = '';
    let steps = [];
    let toolResults = [];

    turnsBadge.style.display = 'inline-flex';
    turnsCount.textContent = '...';

    try {
        // Get file IDs from attached files
        const fileIds = attachedFiles.map(f => f.fileId);

        const response = await fetch('/api/chat/stream', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                message,
                history: chatHistory,
                attachedFiles: fileIds
            })
        });

        const reader = response.body.getReader();
        const decoder = new TextDecoder();
        let buffer = '';

        while (true) {
            const { done, value } = await reader.read();
            if (done) break;

            buffer += decoder.decode(value, { stream: true });
            const lines = buffer.split('\n\n');
            buffer = lines.pop() || '';

            for (const line of lines) {
                if (!line.trim()) continue;
                const nlIdx = line.indexOf('\n');
                if (nlIdx === -1) continue;
                const eventLine = line.substring(0, nlIdx);
                const dataLine = line.substring(nlIdx + 1);
                if (!eventLine.startsWith('event: ') || !dataLine.startsWith('data: ')) continue;
                const event = eventLine.substring(7);
                let data;
                try { data = JSON.parse(dataLine.substring(6)); } catch { continue; }

                if (event === 'text') {
                    // Append text chunk in real-time
                    accumulatedText += data.chunk;
                    streamingText.innerHTML = formatResponse(accumulatedText);
                    chatMessages.scrollTop = chatMessages.scrollHeight;
                } else if (event === 'tool_start') {
                    const icon = TOOL_ICONS[data.tool] || '&#128295;';
                    const toolName = data.tool.replace(/_/g, ' ');
                    const indicator = document.createElement('div');
                    indicator.className = 'tool-indicator executing';
                    indicator.id = `tool-${data.tool}-${Date.now()}`;
                    indicator.innerHTML = `<span class="tool-spinner">&#9881;</span> ${icon} ${toolName}`;
                    streamingTools.appendChild(indicator);
                    chatMessages.scrollTop = chatMessages.scrollHeight;
                } else if (event === 'tool_end') {
                    // Update the last matching indicator to show success/error
                    const indicators = streamingTools.querySelectorAll('.tool-indicator.executing');
                    for (const ind of indicators) {
                        if (ind.innerHTML.includes(data.tool.replace(/_/g, ' '))) {
                            ind.className = `tool-indicator ${data.success ? 'success' : 'error'}`;
                            const statusIcon = data.success ? '&#10003;' : '&#10007;';
                            const icon = TOOL_ICONS[data.tool] || '&#128295;';
                            const toolName = data.tool.replace(/_/g, ' ');
                            ind.innerHTML = `${statusIcon} ${icon} ${toolName}`;
                            break;
                        }
                    }
                } else if (event === 'done') {
                    // Final result received
                    if (data.turnsUsed > 0) {
                        turnsCount.textContent = data.turnsUsed;
                    } else {
                        turnsBadge.style.display = 'none';
                    }

                    // Handle automatic downloads
                    // NOTE: download_drive_file_to_local is a silent intermediate step for email
                    // attachments — do NOT trigger a browser download for it. Only explicit download
                    // tools (download_drive_file, gcs_download_object, conversions) should trigger downloads.
                    const automaticDownloads = (data.toolResults || [])
                        .filter(result =>
                            result &&
                            !result.error &&
                            (result.tool === 'download_drive_file' || result.tool === 'convert_file_to_google_doc' || result.tool === 'convert_file_to_google_sheet' || result.tool === 'gcs_download_object') &&
                            result.result &&
                            result.result.downloadUrl
                        );
                    for (const item of automaticDownloads) {
                        triggerBrowserDownload(item.result.downloadUrl, item.result.downloadName || item.result.name || 'download');
                    }

                    // Clear temporary streaming elements and build final HTML
                    streamingText.innerHTML = '';
                    streamingTools.innerHTML = '';

                    let finalHtml = '';

                    const finalResponseText = String(data.response || '').trim() || String(accumulatedText || '').trim();
                    if (finalResponseText) {
                        finalHtml += formatResponse(finalResponseText);
                    }

                    // Never leave an empty assistant bubble after streaming.
                    if (!finalHtml.trim()) {
                        finalHtml = formatResponse('I completed your request, but no response text was returned. Please try again.');
                    }

                    assistantMsgDiv.querySelector('.message-bubble').innerHTML = finalHtml;
                    pushChatHistoryEntry('user', message);
                    pushChatHistoryEntry('assistant', finalResponseText || '');
                } else if (event === 'error') {
                    streamingText.innerHTML = `<p style="color: #ef4444;">Error: ${escapeHtml(data.error)}</p>`;
                    turnsBadge.style.display = 'none';
                }
            }
        }
    } catch (error) {
        console.error('Stream error:', error);
        streamingText.innerHTML = `<p style="color: #ef4444;">Error: ${error.message || 'Failed to get response'}</p>`;
        turnsBadge.style.display = 'none';
    } finally {
        // Clear attached files after sending
        attachedFiles = [];
        updateAttachedFilesUI();

        sendBtn.disabled = false;
        if (!isMobile()) {
            messageInput.focus();
        }
    }
}

function isMobile() {
    return /Android|webOS|iPhone|iPad|iPod|BlackBerry|IEMobile|Opera Mini/i.test(navigator.userAgent);
}

// Speech Recognition
function setupSpeechRecognition() {
    if (!('webkitSpeechRecognition' in window) && !('SpeechRecognition' in window)) {
        micBtn.style.display = 'none';
        return;
    }

    const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
    recognition = new SpeechRecognition();
    recognition.continuous = false;
    recognition.interimResults = false;
    recognition.lang = 'en-US';

    recognition.onstart = () => {
        isRecording = true;
        micBtn.classList.add('listening');
        micBtn.innerHTML = `
            <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                <rect x="6" y="6" width="12" height="12" rx="2"/>
            </svg>
        `;
        messageInput.placeholder = "Listening...";
    };

    recognition.onend = () => {
        isRecording = false;
        micBtn.classList.remove('listening');
        micBtn.innerHTML = `
            <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                <path d="M12 1a3 3 0 0 0-3 3v8a3 3 0 0 0 6 0V4a3 3 0 0 0-3-3z"/>
                <path d="M19 10v2a7 7 0 0 1-14 0v-2"/>
                <line x1="12" y1="19" x2="12" y2="23"/>
                <line x1="8" y1="23" x2="16" y2="23"/>
            </svg>
        `;
        messageInput.placeholder = "Ask me anything across Gmail, Calendar, Chat, Drive, Sheets, or GitHub...";
    };

    recognition.onresult = (event) => {
        const transcript = event.results[0][0].transcript;
        const currentText = messageInput.value;
        const spacing = currentText && !currentText.endsWith(' ') ? ' ' : '';
        messageInput.value = currentText + spacing + transcript;
        autoResizeTextarea();
        sendBtn.disabled = false;
        messageInput.focus();
    };

    recognition.onerror = (event) => {
        console.error('Speech recognition error', event.error);
        isRecording = false;
        micBtn.classList.remove('listening');
        messageInput.placeholder = "Error accessing microphone";
        setTimeout(() => {
            messageInput.placeholder = "Ask me anything across Gmail, Calendar, Chat, Drive, Sheets, or GitHub...";
        }, 3000);
    };

    micBtn.addEventListener('click', () => {
        if (isRecording) {
            recognition.stop();
        } else {
            try {
                recognition.start();
            } catch (e) {
                console.error('Recognition start error:', e);
            }
        }
    });
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
    if (!text || !String(text).trim()) return '';

    const codeBlocks = [];
    const normalized = String(text).replace(/\r\n?/g, '\n');
    const withCodePlaceholders = normalized.replace(/```([a-zA-Z0-9_-]+)?\n?([\s\S]*?)```/g, (_, language, code) => {
        const index = codeBlocks.length;
        codeBlocks.push({
            language: escapeHtml(String(language || '').trim().toLowerCase()),
            code: escapeHtml(String(code || '').replace(/\n+$/, ''))
        });
        return `@@CODE_BLOCK_${index}@@`;
    });

    const htmlParts = [];
    const lines = withCodePlaceholders.split('\n');
    let listMode = null;

    const closeOpenList = () => {
        if (!listMode) return;
        htmlParts.push(listMode === 'ol' ? '</ol>' : '</ul>');
        listMode = null;
    };

    for (const rawLine of lines) {
        const line = rawLine.trim();

        if (!line) {
            closeOpenList();
            continue;
        }

        const codeMatch = line.match(/^@@CODE_BLOCK_(\d+)@@$/);
        if (codeMatch) {
            closeOpenList();
            const block = codeBlocks[Number.parseInt(codeMatch[1], 10)] || { language: '', code: '' };
            const languageClass = block.language ? ` class="language-${block.language}"` : '';
            htmlParts.push(`<pre class="assistant-code"><code${languageClass}>${block.code}</code></pre>`);
            continue;
        }

        const headingMatch = line.match(/^(#{1,3})\s+(.+)$/);
        if (headingMatch) {
            closeOpenList();
            const level = Math.min(3, headingMatch[1].length);
            htmlParts.push(`<h${level}>${applyInlineFormatting(headingMatch[2])}</h${level}>`);
            continue;
        }

        if (/^---+$/.test(line)) {
            closeOpenList();
            htmlParts.push('<hr>');
            continue;
        }

        const unorderedMatch = line.match(/^[-*]\s+(.+)$/);
        if (unorderedMatch) {
            if (listMode !== 'ul') {
                closeOpenList();
                htmlParts.push('<ul>');
                listMode = 'ul';
            }
            htmlParts.push(`<li>${applyInlineFormatting(unorderedMatch[1])}</li>`);
            continue;
        }

        const orderedMatch = line.match(/^\d+\.\s+(.+)$/);
        if (orderedMatch) {
            if (listMode !== 'ol') {
                closeOpenList();
                htmlParts.push('<ol>');
                listMode = 'ol';
            }
            htmlParts.push(`<li>${applyInlineFormatting(orderedMatch[1])}</li>`);
            continue;
        }

        const quoteMatch = line.match(/^>\s?(.+)$/);
        if (quoteMatch) {
            closeOpenList();
            htmlParts.push(`<blockquote>${applyInlineFormatting(quoteMatch[1])}</blockquote>`);
            continue;
        }

        closeOpenList();
        htmlParts.push(`<p>${applyInlineFormatting(line)}</p>`);
    }

    closeOpenList();
    return `<div class="assistant-response">${htmlParts.join('')}</div>`;
}

function applyInlineFormatting(text) {
    let safe = escapeHtml(String(text || ''));
    safe = safe.replace(/\[([^\]]+)\]\((https?:\/\/[^\s)]+)\)/g, '<a href="$2" target="_blank" rel="noopener noreferrer">$1</a>');
    safe = safe.replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>');
    safe = safe.replace(/`([^`]+)`/g, '<code>$1</code>');
    safe = safe.replace(/\*([^*\n]+)\*/g, '<em>$1</em>');
    return safe;
}

// Format tool results
function formatToolResults(results) {
    const truncateText = (value, max = 500) => {
        const text = String(value || '');
        if (text.length <= max) return text;
        return `${text.slice(0, max)}...`;
    };

    const formatPre = (value, extraClass = '') => {
        const className = ['result-pre', extraClass].filter(Boolean).join(' ');
        return `<pre class="${className}">${escapeHtml(String(value || ''))}</pre>`;
    };

    const actionLink = ({ href, label, variant = 'primary', download }) => {
        if (!href || !label) return '';
        const linkClass = variant === 'secondary' ? 'result-link result-link-secondary' : 'result-link result-link-primary';
        const downloadAttr = download ? ` download="${escapeHtml(download)}"` : '';
        return `<a href="${escapeHtml(href)}"${downloadAttr} target="_blank" rel="noopener noreferrer" class="${linkClass}">${escapeHtml(label)}</a>`;
    };

    const actionRow = (...links) => {
        const valid = links.filter(Boolean);
        if (!valid.length) return '';
        return `<div class="result-actions">${valid.join('')}</div>`;
    };

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
                    <div class="email-card email-clickable" data-message-id="${escapeHtml(email.id || '')}">
                        <div class="email-card-content-wrapper">
                            <div class="email-avatar">${escapeHtml(initials)}</div>
                            <div class="email-main-info">
                                <div class="email-sender">${escapeHtml(displayName)}</div>
                                <div class="email-snippet-wrapper">
                                    <span class="email-subject">${escapeHtml(email.subject || '(no subject)')}</span>
                                    <span class="email-snippet">- ${escapeHtml(email.snippet ? email.snippet.slice(0, 300) : '')}...</span>
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
                        <p class="result-body-preview">${escapeHtml(truncateText(result.result.body || '', 500))}</p>
                    </div>
                `;
                // Labels
            } else if (result.result.labels && Array.isArray(result.result.labels)) {
                content = `<div class="labels-list">${result.result.labels.map(l => `<span class="label-chip">${escapeHtml(l.name)}</span>`).join('')}</div>`;
                // Drafts
            } else if (result.result.drafts && Array.isArray(result.result.drafts)) {
                content = result.result.drafts.map(d => `
                    <div class="email-card vertical">
                        <div class="email-card-header">
                            <span class="email-card-subject">${escapeHtml(d.subject || '(no subject)')}</span>
                        </div>
                        <div class="email-card-from">To: ${escapeHtml(d.to || 'Not set')}</div>
                    </div>
                `).join('');
                // Google Chat spaces
            } else if (result.result.spaces && Array.isArray(result.result.spaces)) {
                content = result.result.spaces.map(s => `
                    <div class="email-card vertical">
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
                    <div class="email-card thread-msg vertical">
                        <div class="email-card-header">
                            <span class="email-card-from result-strong">${escapeHtml(msg.sender || '')}</span>
                            <span class="email-card-date">${formatDate(msg.createTime)}</span>
                        </div>
                        <div class="email-card-snippet">${escapeHtml(msg.text || '')}</div>
                    </div>
                `).join('');
                // Thread messages
            } else if (result.result.messages && Array.isArray(result.result.messages)) {
                content = result.result.messages.map(msg => `
                    <div class="email-card thread-msg vertical">
                        <div class="email-card-header">
                            <span class="email-card-from result-strong">${escapeHtml(msg.from || '')}</span>
                            <span class="email-card-date">${formatDate(msg.date)}</span>
                        </div>
                        <p class="result-line-preview">${escapeHtml(truncateText(msg.body || msg.snippet || '', 200))}</p>
                    </div>
                `).join('');
                // Attachments
            } else if (result.result.attachments && Array.isArray(result.result.attachments)) {
                content = result.result.attachments.map(a => `
                    <div class="attachment-card">
                        <span class="attachment-icon">&#128206;</span>
                        <div class="attachment-meta">
                            <div class="attachment-title">${escapeHtml(a.filename || a.name || 'attachment')}</div>
                            <div class="attachment-subtext">${escapeHtml(a.mimeType || a.contentType || 'unknown')} &middot; ${formatSize(a.size)}</div>
                        </div>
                    </div>
                `).join('');
                // Calendar events
            } else if (result.result.events && Array.isArray(result.result.events)) {
                content = result.result.events.map(e => `
                    <div class="email-card vertical">
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
                    <div class="email-card vertical">
                        <div class="email-card-header">
                            <span class="email-card-subject">${escapeHtml(c.summary || c.id)}</span>
                            ${c.primary ? '<span class="result-pill">Primary</span>' : ''}
                        </div>
                        ${c.description ? `<div class="email-card-snippet">${escapeHtml(c.description)}</div>` : ''}
                    </div>
                `).join('');
                // Drive files
            } else if (result.result.files && Array.isArray(result.result.files)) {
                content = result.result.files.map(f => `
                    <div class="email-card vertical">
                        <div class="email-card-header">
                            <span class="email-card-subject">${escapeHtml(f.name || f.id)}</span>
                            <span class="email-card-date">${escapeHtml((f.mimeType || '').replace('application/vnd.google-apps.', ''))}</span>
                        </div>
                        <div class="email-card-from"><code>${escapeHtml(f.id || '')}</code></div>
                        ${f.webViewLink ? `<div class="email-card-snippet"><a href="${escapeHtml(f.webViewLink)}" target="_blank" rel="noopener noreferrer">Open in Drive</a></div>` : ''}
                    </div>
                `).join('');
                // Drive download payload
            } else if ((result.tool === 'download_drive_file' || result.tool === 'download_drive_file_to_local') && result.result.downloadUrl) {
                const downloadHref = result.result.downloadUrl;
                const downloadName = result.result.downloadName || result.result.name || 'download';
                const formatLabel = escapeHtml(result.result.format || (result.tool === 'download_drive_file_to_local' ? 'local' : 'raw'));
                content = `
                    <div class="email-card vertical">
                        <div class="email-card-header">
                            <span class="email-card-subject">${escapeHtml(result.result.name || 'Drive file')}</span>
                            <span class="email-card-date">${formatLabel.toUpperCase()}</span>
                        </div>
                        <div class="email-card-from"><code>${escapeHtml(result.result.fileId || '')}</code></div>
                        ${actionRow(
                            actionLink({ href: downloadHref, label: 'Download File', variant: 'primary', download: downloadName }),
                            result.result.webViewLink ? actionLink({ href: result.result.webViewLink, label: 'Open in Drive', variant: 'secondary' }) : ''
                        )}
                    </div>
                `;
                // GCS download payload
            } else if (result.tool === 'gcs_download_object' && result.result.downloadUrl) {
                const downloadHref = result.result.downloadUrl;
                const downloadName = result.result.downloadName || result.result.name || 'download';
                const sizeLabel = result.result.size ? `${(Number(result.result.size) / 1024).toFixed(1)} KB` : '';
                content = `
                    <div class="email-card vertical">
                        <div class="email-card-header">
                            <span class="email-card-subject">${escapeHtml(result.result.name || 'GCS object')}</span>
                            ${sizeLabel ? `<span class="email-card-date">${escapeHtml(sizeLabel)}</span>` : ''}
                        </div>
                        <div class="email-card-from"><code>${escapeHtml(result.result.bucket || '')}</code></div>
                        ${actionRow(
                            actionLink({ href: downloadHref, label: 'Download File', variant: 'primary', download: downloadName })
                        )}
                    </div>
                `;
                // Extracted Drive text
            } else if (result.tool === 'extract_drive_file_text' && result.result.content) {
                const preview = String(result.result.content || '');
                content = `
                    <div class="email-card vertical">
                        <div class="email-card-header">
                            <span class="email-card-subject">${escapeHtml(result.result.name || 'Extracted text')}</span>
                            <span class="email-card-date">${escapeHtml(result.result.extractionMethod || '')}</span>
                        </div>
                        ${formatPre(preview, 'result-pre-tall')}
                    </div>
                `;
                // Appended text to Drive document
            } else if (result.tool === 'append_drive_document_text' && result.result.fileId) {
                content = `
                    <div class="email-card vertical">
                        <div class="email-card-header">
                            <span class="email-card-subject">${escapeHtml(result.result.name || 'Drive document')}</span>
                            <span class="email-card-date">${result.result.usedDocsApi ? 'Docs API' : 'Drive Fallback'}</span>
                        </div>
                        <div class="email-card-from"><code>${escapeHtml(result.result.fileId || '')}</code></div>
                        <div class="email-card-snippet result-note">
                            ${result.result.message ? escapeHtml(result.result.message) : 'Text appended successfully.'}
                        </div>
                        ${result.result.webViewLink ? actionRow(actionLink({ href: result.result.webViewLink, label: 'Open in Drive', variant: 'primary' })) : ''}
                    </div>
                `;
                // Converted file to Google Doc
            } else if (result.tool === 'convert_file_to_google_doc' && result.result.documentId) {
                content = `
                    <div class="email-card vertical">
                        <div class="email-card-header">
                            <span class="email-card-subject">${escapeHtml(result.result.name || 'Converted Google Doc')}</span>
                            <span class="email-card-date">Google Doc</span>
                        </div>
                        <div class="email-card-from">Source: ${escapeHtml(result.result.sourceName || result.result.sourceFileId || '')}</div>
                        ${actionRow(
                            result.result.webViewLink ? actionLink({ href: result.result.webViewLink, label: 'Open Converted Doc', variant: 'primary' }) : '',
                            result.result.downloadUrl
                                ? actionLink({
                                    href: result.result.downloadUrl,
                                    label: 'Download Converted File',
                                    variant: 'secondary',
                                    download: result.result.downloadName || result.result.name || 'converted-document'
                                })
                                : ''
                        )}
                    </div>
                `;
                // Converted file to Google Sheet
            } else if (result.tool === 'convert_file_to_google_sheet' && result.result.spreadsheetId) {
                content = `
                    <div class="email-card vertical">
                        <div class="email-card-header">
                            <span class="email-card-subject">${escapeHtml(result.result.name || 'Converted Google Sheet')}</span>
                            <span class="email-card-date">Google Sheet</span>
                        </div>
                        <div class="email-card-from">Source: ${escapeHtml(result.result.sourceName || result.result.sourceFileId || '')}</div>
                        ${actionRow(
                            result.result.webViewLink ? actionLink({ href: result.result.webViewLink, label: 'Open Converted Sheet', variant: 'primary' }) : '',
                            result.result.downloadUrl
                                ? actionLink({
                                    href: result.result.downloadUrl,
                                    label: 'Download Converted File',
                                    variant: 'secondary',
                                    download: result.result.downloadName || result.result.name || 'converted-sheet'
                                })
                                : ''
                        )}
                    </div>
                `;
                // Spreadsheet list
            } else if (result.result.spreadsheets && Array.isArray(result.result.spreadsheets)) {
                content = result.result.spreadsheets.map(s => `
                    <div class="email-card vertical">
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
                content = formatPre(JSON.stringify(preview, null, 2), 'result-pre-compact');
                // GitHub repos
            } else if (result.result.repos && Array.isArray(result.result.repos)) {
                content = result.result.repos.map(r => `
                    <div class="email-card vertical">
                        <div class="email-card-header">
                            <span class="email-card-subject">${escapeHtml(r.full_name || r.name)}</span>
                            <span class="email-card-date">${r.language || ''}</span>
                        </div>
                        ${r.description ? `<div class="email-card-snippet">${escapeHtml(r.description.slice(0, 100))}</div>` : ''}
                        <div class="email-card-from result-metrics">&#11088; ${r.stars || 0} &middot; &#128204; ${r.forks || 0}${r.private ? ' &middot; Private' : ''}</div>
                        ${r.url ? actionRow(actionLink({ href: r.url, label: 'Go to Repo', variant: 'primary' })) : ''}
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
                        ${i.labels && i.labels.length > 0 ? `<div class="labels-list compact">${i.labels.map(l => `<span class="label-chip">${escapeHtml(l)}</span>`).join('')}</div>` : ''}
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
                            <code class="result-code-highlight">${escapeHtml(c.sha)}</code>
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
                content = formatPre(typeof summary === 'string' ? summary : JSON.stringify(summary, null, 2), 'result-pre-compact');
            }
        } else if (result.error) {
            content = `<p class="result-error-text">${escapeHtml(result.error)}</p>`;
        }

        return `
            <div class="tool-result ${isError ? 'tool-result-error' : ''}">
                <div class="tool-result-header">
                    <span class="tool-result-icon">${toolIcon}</span>
                    <span class="tool-result-title">${title}</span>
                    <span class="tool-result-status ${isError ? 'is-error' : 'is-success'}">${icon}</span>
                </div>
                ${content}
            </div>
        `;
    }).join('');
}

// Open email when card is clicked
async function openEmail(messageId, cardElement) {
    if (!messageId || !cardElement) return;
    // Security: validate messageId format
    if (!/^[a-zA-Z0-9_-]+$/.test(messageId)) return;

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
                
                <div class="email-details-body">${data.bodyHtml ? sanitizeHtml(data.bodyHtml) : escapeHtml(data.body || '(No content)')}</div>
                
                ${attachmentsHtml}

                <div class="email-actions" style="margin-top: 1.5rem; display: flex; gap: 0.5rem;">
                    <button class="action-btn email-reply-btn" data-msg-id="${escapeHtml(messageId)}" style="padding:0.5rem 1rem; background:var(--accent-primary); color:white; border:none; border-radius:6px; cursor:pointer">Reply</button>
                    <button class="action-btn email-forward-btn" data-msg-id="${escapeHtml(messageId)}" style="padding:0.5rem 1rem; background:var(--bg-tertiary); color:var(--text-primary); border:1px solid var(--border-color); border-radius:6px; cursor:pointer">Forward</button>
                </div>
            </div>
        `;

        // Attach event listeners safely (no inline onclick)
        const replyBtn = detailsDiv.querySelector('.email-reply-btn');
        const forwardBtn = detailsDiv.querySelector('.email-forward-btn');
        if (replyBtn) replyBtn.addEventListener('click', (e) => { e.stopPropagation(); messageInput.value = `Reply to email ${replyBtn.dataset.msgId} saying...`; messageInput.focus(); });
        if (forwardBtn) forwardBtn.addEventListener('click', (e) => { e.stopPropagation(); messageInput.value = `Forward email ${forwardBtn.dataset.msgId} to...`; messageInput.focus(); });
    } catch (error) {
        detailsDiv.innerHTML = `<div style="color:var(--error); padding:1rem">Error loading email</div>`;
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
    const bubbleClass = role === 'assistant' && allowHtml
        ? 'message-bubble assistant-rich-bubble'
        : 'message-bubble';

    messageDiv.innerHTML = `
        <div class="message-avatar">${avatar}</div>
        <div class="message-content">
            <div class="${bubbleClass}">${bubbleContent}</div>
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

// Security: sanitize HTML from email bodies - strip scripts, event handlers, iframes
function sanitizeHtml(html) {
    if (!html) return '';
    const doc = new DOMParser().parseFromString(html, 'text/html');
    // Remove dangerous elements
    doc.querySelectorAll('script, iframe, object, embed, form, base, meta, link').forEach(el => el.remove());
    // Remove all event handler attributes (onclick, onerror, onload, etc.)
    doc.querySelectorAll('*').forEach(el => {
        for (const attr of [...el.attributes]) {
            if (attr.name.startsWith('on') || attr.name === 'srcdoc') el.removeAttribute(attr.name);
            if (attr.name === 'href' && attr.value.trim().toLowerCase().startsWith('javascript:')) el.removeAttribute(attr.name);
            if (attr.name === 'src' && attr.value.trim().toLowerCase().startsWith('javascript:')) el.removeAttribute(attr.name);
        }
    });
    return doc.body.innerHTML;
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
