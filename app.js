/**
 * SheetSDK Web Application — Main orchestration
 */

const CLIENT_ID = '482589965597-7n9m7o2u524ngooo8pna4lgqrf2bp970.apps.googleusercontent.com';

// State
let currentFile = null;
let parsedSpec = null;
let isSignedIn = false;

// DOM refs
const dropZone = document.getElementById('drop-zone');
const fileInput = document.getElementById('file-input');
const fileInfo = document.getElementById('file-info');
const fileName = document.getElementById('file-name');
const fileStatus = document.getElementById('file-status');
const previewSection = document.getElementById('preview-section');
const previewBody = document.getElementById('preview-body');
const createBtn = document.getElementById('create-btn');
const createSection = document.getElementById('create-section');
const statusSection = document.getElementById('status-section');
const statusMessage = document.getElementById('status-message');
const resultSection = document.getElementById('result-section');
const resultLink = document.getElementById('result-link');
const signInBtn = document.getElementById('sign-in-btn');
const signOutBtn = document.getElementById('sign-out-btn');
const authStatus = document.getElementById('auth-status');
const titleInput = document.getElementById('spreadsheet-title');
const removeFileBtn = document.getElementById('remove-file');

// Initialize Google Auth
function initApp() {
    initGoogleAuth(CLIENT_ID, handleSignIn, handleSignOut);
    setupFileHandlers();
    updateUI();
}

function handleSignIn(tokenResponse) {
    isSignedIn = true;
    updateUI();
}

function handleSignOut() {
    isSignedIn = false;
    updateUI();
}

function setupFileHandlers() {
    // Click to browse
    dropZone.addEventListener('click', (e) => {
        if (e.target.closest('#remove-file')) return;
        fileInput.click();
    });

    fileInput.addEventListener('change', (e) => {
        if (e.target.files.length > 0) {
            handleFile(e.target.files[0]);
        }
    });

    // Drag and drop
    dropZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropZone.classList.add('drag-over');
    });

    dropZone.addEventListener('dragleave', () => {
        dropZone.classList.remove('drag-over');
    });

    dropZone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropZone.classList.remove('drag-over');
        if (e.dataTransfer.files.length > 0) {
            handleFile(e.dataTransfer.files[0]);
        }
    });

    // Remove file
    removeFileBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        clearFile();
    });

    // Create button
    createBtn.addEventListener('click', handleCreate);

    // Sign in/out
    signInBtn.addEventListener('click', () => requestSignIn());
    signOutBtn.addEventListener('click', () => signOut(handleSignOut));
}

function handleFile(file) {
    if (!file.name.endsWith('.js')) {
        showFileError('Please select a .js file');
        return;
    }

    const reader = new FileReader();
    reader.onload = (e) => {
        const content = e.target.result;

        if (!validateHeader(content)) {
            showFileError('Invalid file: must start with "use sheetsdk"');
            currentFile = null;
            parsedSpec = null;
            updateUI();
            return;
        }

        try {
            parsedSpec = parseSheetSDK(content);
            currentFile = { name: file.name, content };
            showPreview();
            updateUI();
        } catch (err) {
            showFileError(`Parse error: ${err.message}`);
            currentFile = null;
            parsedSpec = null;
            updateUI();
        }
    };
    reader.readAsText(file);
}

function showFileError(msg) {
    fileInfo.classList.remove('hidden');
    fileName.textContent = msg;
    fileStatus.textContent = '✗';
    fileStatus.className = 'file-status error';
    removeFileBtn.classList.remove('hidden');
}

function clearFile() {
    currentFile = null;
    parsedSpec = null;
    fileInput.value = '';
    fileInfo.classList.add('hidden');
    previewSection.classList.add('hidden');
    resultSection.classList.add('hidden');
    statusSection.classList.add('hidden');
    updateUI();
}

function showPreview() {
    fileInfo.classList.remove('hidden');
    fileName.textContent = currentFile.name;
    fileStatus.textContent = '✓';
    fileStatus.className = 'file-status success';
    removeFileBtn.classList.remove('hidden');

    // Build preview
    previewBody.innerHTML = '';

    for (const sheet of parsedSpec.sheets) {
        const row = document.createElement('tr');
        row.innerHTML = `
      <td>
        <span class="sheet-name">${escapeHtml(sheet.name)}</span>
        ${sheet.isPrimary ? '<span class="badge primary">primary</span>' : ''}
      </td>
      <td>${sheet.formulas.length}</td>
      <td>${sheet.values.length}</td>
      <td>${sheet.notes.length}</td>
      <td>${sheet.conditionalHighlights.length}</td>
    `;
        previewBody.appendChild(row);
    }

    previewSection.classList.remove('hidden');

    // Auto-title from filename
    if (!titleInput.value) {
        titleInput.value = currentFile.name.replace(/\.js$/, '');
    }
}

function escapeHtml(str) {
    const div = document.createElement('div');
    div.textContent = str;
    return div.innerHTML;
}

async function handleCreate() {
    if (!isSignedIn) {
        showStatus('Please sign in with Google first', 'error');
        return;
    }
    if (!parsedSpec) {
        showStatus('Please upload a valid .js file first', 'error');
        return;
    }

    const title = titleInput.value.trim() || 'Untitled SheetSDK Spreadsheet';

    createBtn.disabled = true;
    createBtn.innerHTML = '<span class="spinner"></span> Creating...';
    showStatus('Creating spreadsheet...', 'loading');
    resultSection.classList.add('hidden');

    try {
        const scriptCode = extractScriptCode(currentFile.content);
        const result = await createSpreadsheet(parsedSpec, title, scriptCode || null);
        showStatus('Spreadsheet created successfully!', 'success');
        resultLink.href = result.url;
        resultLink.textContent = title;
        resultSection.classList.remove('hidden');
    } catch (err) {
        showStatus(`Error: ${err.message}`, 'error');
    } finally {
        createBtn.disabled = false;
        createBtn.innerHTML = '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="12" y1="11" x2="12" y2="17"/><line x1="9" y1="14" x2="15" y2="14"/></svg> Create Spreadsheet';
    }
}

function showStatus(msg, type) {
    statusSection.classList.remove('hidden');
    statusMessage.textContent = msg;
    statusMessage.className = `status-msg ${type}`;
}

function updateUI() {
    // Auth state
    if (isSignedIn) {
        signInBtn.classList.add('hidden');
        signOutBtn.classList.remove('hidden');
        authStatus.textContent = 'Signed in';
        authStatus.className = 'auth-badge signed-in';
    } else {
        signInBtn.classList.remove('hidden');
        signOutBtn.classList.add('hidden');
        authStatus.textContent = 'Not signed in';
        authStatus.className = 'auth-badge signed-out';
    }

    // Create section visibility
    if (parsedSpec && isSignedIn) {
        createSection.classList.remove('hidden');
    } else if (parsedSpec) {
        createSection.classList.remove('hidden');
    } else {
        createSection.classList.add('hidden');
    }
}

// Boot
window.addEventListener('load', initApp);
