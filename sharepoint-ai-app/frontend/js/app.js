/**
 * SharePoint AI - Frontend Application
 */

// ============================================
// Estado da Aplicação
// ============================================

const AppState = {
    authenticated: false,
    user: null,
    documents: [],
    selectedDocuments: [],
    summaries: []
};

// ============================================
// Inicialização
// ============================================

document.addEventListener('DOMContentLoaded', () => {
    checkAuthStatus();
    setupEventListeners();
    checkAuthCallback();
});

// ============================================
// Event Listeners
// ============================================

function setupEventListeners() {
    // Login/Logout
    document.getElementById('loginBtn').addEventListener('click', handleLogin);
    document.getElementById('logoutBtn').addEventListener('click', handleLogout);

    // Documents
    document.getElementById('refreshBtn').addEventListener('click', loadDocuments);
    document.getElementById('selectAll').addEventListener('click', handleSelectAll);
    document.getElementById('summarizeBtn').addEventListener('click', handleSummarize);
}

// ============================================
// Autenticação
// ============================================

async function checkAuthStatus() {
    try {
        const response = await fetch('/api/auth/status', {
            credentials: 'include'
        });
        const data = await response.json();

        if (data.authenticated) {
            AppState.authenticated = true;
            AppState.user = {
                name: data.user_name,
                email: data.user_email
            };
            showDocumentsSection();
            loadDocuments();
        } else {
            showLoginSection();
        }
    } catch (error) {
        console.error('Erro ao verificar autenticação:', error);
        showLoginSection();
    }
}

function checkAuthCallback() {
    const urlParams = new URLSearchParams(window.location.search);

    if (urlParams.get('authenticated') === 'true') {
        showToast('Login realizado com sucesso!', 'success');
        window.history.replaceState({}, '', '/');
    }

    if (urlParams.get('error')) {
        const error = urlParams.get('error');
        showToast(`Erro na autenticação: ${error}`, 'error');
        window.history.replaceState({}, '', '/');
    }
}

async function handleLogin() {
    try {
        const response = await fetch('/api/auth/login');
        const data = await response.json();

        if (data.auth_url) {
            // Redireciona para página de login da Microsoft
            window.location.href = data.auth_url;
        } else {
            showToast('Erro ao iniciar login', 'error');
        }
    } catch (error) {
        console.error('Erro no login:', error);
        showToast('Erro ao conectar com o servidor', 'error');
    }
}

async function handleLogout() {
    try {
        await fetch('/api/auth/logout');
        AppState.authenticated = false;
        AppState.user = null;
        AppState.documents = [];
        AppState.selectedDocuments = [];
        showLoginSection();
        showToast('Logout realizado com sucesso', 'success');
    } catch (error) {
        console.error('Erro no logout:', error);
        showToast('Erro ao fazer logout', 'error');
    }
}

// ============================================
// UI - Seções
// ============================================

function showLoginSection() {
    document.getElementById('loginSection').style.display = 'block';
    document.getElementById('documentsSection').style.display = 'none';
    document.getElementById('userInfo').style.display = 'none';
}

function showDocumentsSection() {
    document.getElementById('loginSection').style.display = 'none';
    document.getElementById('documentsSection').style.display = 'block';
    document.getElementById('userInfo').style.display = 'flex';
    document.getElementById('userName').textContent = AppState.user.name;
}

// ============================================
// Documentos
// ============================================

async function loadDocuments() {
    const loadingEl = document.getElementById('documentsLoading');
    const tableEl = document.getElementById('documentsTable');
    const noDocsEl = document.getElementById('noDocuments');

    // Mostra loading
    loadingEl.style.display = 'block';
    tableEl.style.display = 'none';
    noDocsEl.style.display = 'none';

    try {
        const response = await fetch('/api/sharepoint/documents', {
            credentials: 'include'
        });

        if (!response.ok) {
            throw new Error('Erro ao carregar documentos');
        }

        const data = await response.json();
        AppState.documents = data.documents;

        loadingEl.style.display = 'none';

        if (data.documents.length === 0) {
            noDocsEl.style.display = 'block';
        } else {
            tableEl.style.display = 'block';
            renderDocumentsTable(data.documents);
            showToast(`${data.documents.length} documentos carregados`, 'success');
        }
    } catch (error) {
        console.error('Erro ao carregar documentos:', error);
        loadingEl.style.display = 'none';
        showToast('Erro ao carregar documentos do SharePoint', 'error');
    }
}

function renderDocumentsTable(documents) {
    const tbody = document.getElementById('documentsTableBody');
    tbody.innerHTML = '';

    documents.forEach(doc => {
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>
                <input type="checkbox" class="doc-checkbox" data-doc-id="${doc.id}"
                       data-doc-name="${doc.name}"
                       data-doc-type="${doc.type}"
                       data-drive-id="${doc.driveId}">
            </td>
            <td>
                <strong>${doc.name}</strong>
            </td>
            <td>
                <span class="file-type-badge badge-${doc.type}">
                    ${getFileTypeLabel(doc.type)}
                </span>
            </td>
            <td>${formatFileSize(doc.size)}</td>
            <td>${formatDate(doc.lastModified)}</td>
            <td>
                <a href="${doc.webUrl}" target="_blank" class="btn btn-sm btn-secondary">
                    Abrir
                </a>
            </td>
        `;
        tbody.appendChild(row);
    });

    // Adiciona event listeners aos checkboxes
    document.querySelectorAll('.doc-checkbox').forEach(checkbox => {
        checkbox.addEventListener('change', handleDocumentSelection);
    });
}

function handleSelectAll(e) {
    const checkboxes = document.querySelectorAll('.doc-checkbox');
    checkboxes.forEach(cb => {
        cb.checked = e.target.checked;
    });
    updateSelectedDocuments();
}

function handleDocumentSelection() {
    updateSelectedDocuments();
}

function updateSelectedDocuments() {
    const checkboxes = document.querySelectorAll('.doc-checkbox:checked');
    AppState.selectedDocuments = Array.from(checkboxes).map(cb => ({
        id: cb.dataset.docId,
        name: cb.dataset.docName,
        type: cb.dataset.docType,
        driveId: cb.dataset.driveId
    }));

    // Atualiza UI
    document.getElementById('selectedCount').textContent =
        `${AppState.selectedDocuments.length} documentos selecionados`;

    document.getElementById('summarizeBtn').disabled =
        AppState.selectedDocuments.length === 0;
}

// ============================================
// Resumos com IA
// ============================================

async function handleSummarize() {
    if (AppState.selectedDocuments.length === 0) {
        showToast('Selecione pelo menos um documento', 'error');
        return;
    }

    const summariesSection = document.getElementById('summariesSection');
    const summariesLoading = document.getElementById('summariesLoading');
    const summariesContent = document.getElementById('summariesContent');

    // Mostra seção de resumos
    summariesSection.style.display = 'block';
    summariesLoading.style.display = 'block';
    summariesContent.innerHTML = '';

    // Scroll suave para seção de resumos
    summariesSection.scrollIntoView({ behavior: 'smooth', block: 'start' });

    try {
        showToast('Processando documentos com IA...', 'info');

        const response = await fetch('/api/ai/summarize', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            credentials: 'include',
            body: JSON.stringify({
                document_ids: AppState.selectedDocuments
            })
        });

        if (!response.ok) {
            throw new Error('Erro ao processar documentos');
        }

        const data = await response.json();
        summariesLoading.style.display = 'none';

        renderSummaries(data.summaries, data.consolidated_summary);
        showToast('Resumos gerados com sucesso!', 'success');

    } catch (error) {
        console.error('Erro ao gerar resumos:', error);
        summariesLoading.style.display = 'none';
        showToast('Erro ao gerar resumos. Tente novamente.', 'error');
    }
}

function renderSummaries(summaries, consolidatedSummary) {
    const container = document.getElementById('summariesContent');
    container.innerHTML = '';

    // Resumo consolidado (se múltiplos documentos)
    if (consolidatedSummary && summaries.length > 1) {
        const consolidatedDiv = document.createElement('div');
        consolidatedDiv.className = 'summary-item summary-consolidated';
        consolidatedDiv.innerHTML = `
            <div class="summary-header">
                <div class="summary-title">
                    <span>📊</span>
                    <span>Resumo Consolidado (${summaries.length} documentos)</span>
                </div>
            </div>
            <div class="summary-content">${consolidatedSummary}</div>
        `;
        container.appendChild(consolidatedDiv);
    }

    // Resumos individuais
    summaries.forEach(summary => {
        const summaryDiv = document.createElement('div');
        summaryDiv.className = summary.success ? 'summary-item' : 'summary-item summary-error';

        if (summary.success) {
            summaryDiv.innerHTML = `
                <div class="summary-header">
                    <div class="summary-title">
                        <span>${getFileIcon(summary.file_type)}</span>
                        <span>${summary.file_name}</span>
                    </div>
                    <span class="file-type-badge badge-${summary.file_type}">
                        ${getFileTypeLabel(summary.file_type)}
                    </span>
                </div>
                <div class="summary-content">${summary.summary}</div>
            `;
        } else {
            summaryDiv.innerHTML = `
                <div class="summary-header">
                    <div class="summary-title">
                        <span>❌</span>
                        <span>${summary.file_name}</span>
                    </div>
                </div>
                <div class="error-message">${summary.error}</div>
            `;
        }

        container.appendChild(summaryDiv);
    });
}

// ============================================
// Utilitários
// ============================================

function formatFileSize(bytes) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return Math.round(bytes / Math.pow(k, i) * 100) / 100 + ' ' + sizes[i];
}

function formatDate(dateString) {
    const date = new Date(dateString);
    return date.toLocaleDateString('pt-BR', {
        year: 'numeric',
        month: '2-digit',
        day: '2-digit',
        hour: '2-digit',
        minute: '2-digit'
    });
}

function getFileTypeLabel(type) {
    const labels = {
        'pdf': 'PDF',
        'word': 'Word',
        'excel': 'Excel',
        'powerpoint': 'PowerPoint',
        'text': 'Texto'
    };
    return labels[type] || type.toUpperCase();
}

function getFileIcon(type) {
    const icons = {
        'pdf': '📕',
        'word': '📘',
        'excel': '📗',
        'powerpoint': '📙',
        'text': '📄'
    };
    return icons[type] || '📄';
}

function showToast(message, type = 'info') {
    const container = document.getElementById('toastContainer');
    const toast = document.createElement('div');
    toast.className = `toast toast-${type}`;

    const icons = {
        'success': '✓',
        'error': '✗',
        'info': 'ℹ'
    };

    toast.innerHTML = `
        <span style="font-size: 1.5rem;">${icons[type] || 'ℹ'}</span>
        <span>${message}</span>
    `;

    container.appendChild(toast);

    // Remove toast após 5 segundos
    setTimeout(() => {
        toast.style.animation = 'slideIn 0.3s ease reverse';
        setTimeout(() => toast.remove(), 300);
    }, 5000);
}
