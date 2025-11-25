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
    summaries: [],
    currentPath: '',
    navigationHistory: []
};

// ============================================
// Cache Local (LocalStorage)
// ============================================

const CacheManager = {
    // Tempo de expiração do cache (5 minutos)
    CACHE_DURATION: 5 * 60 * 1000,

    /**
     * Salva dados no cache local
     */
    set(key, data) {
        try {
            const cacheEntry = {
                data: data,
                timestamp: Date.now()
            };
            localStorage.setItem(`cache_${key}`, JSON.stringify(cacheEntry));
        } catch (e) {
            console.warn('Erro ao salvar no cache:', e);
        }
    },

    /**
     * Obtém dados do cache local
     */
    get(key) {
        try {
            const cached = localStorage.getItem(`cache_${key}`);
            if (!cached) return null;

            const cacheEntry = JSON.parse(cached);
            const age = Date.now() - cacheEntry.timestamp;

            // Verifica se o cache expirou
            if (age > this.CACHE_DURATION) {
                this.remove(key);
                return null;
            }

            console.log(`✓ Cache local hit: ${key} (${Math.round(age / 1000)}s atrás)`);
            return cacheEntry.data;
        } catch (e) {
            console.warn('Erro ao ler cache:', e);
            return null;
        }
    },

    /**
     * Remove item do cache
     */
    remove(key) {
        try {
            localStorage.removeItem(`cache_${key}`);
        } catch (e) {
            console.warn('Erro ao remover cache:', e);
        }
    },

    /**
     * Limpa todo o cache
     */
    clear() {
        try {
            const keys = Object.keys(localStorage);
            keys.forEach(key => {
                if (key.startsWith('cache_')) {
                    localStorage.removeItem(key);
                }
            });
            console.log('✓ Cache local limpo');
        } catch (e) {
            console.warn('Erro ao limpar cache:', e);
        }
    }
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
    document.getElementById('refreshBtn').addEventListener('click', () => {
        // Força refresh do cache
        loadDocuments(AppState.currentPath, true);
    });
    document.getElementById('selectAll').addEventListener('click', handleSelectAll);
    document.getElementById('summarizeBtn').addEventListener('click', handleSummarize);

    // Search
    document.getElementById('searchInput').addEventListener('input', handleSearch);
    document.getElementById('clearSearchBtn').addEventListener('click', clearSearch);
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

        // Limpa cache local
        CacheManager.clear();

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

async function loadDocuments(folderPath = null, forceRefresh = false) {
    const loadingEl = document.getElementById('documentsLoading');
    const tableEl = document.getElementById('documentsTable');
    const noDocsEl = document.getElementById('noDocuments');

    // Chave de cache para esta pasta
    const cacheKey = `documents_${folderPath || 'root'}`;

    // Tenta obter do cache local se não for refresh forçado
    if (!forceRefresh) {
        const cachedData = CacheManager.get(cacheKey);
        if (cachedData) {
            AppState.documents = cachedData.documents;
            AppState.currentPath = cachedData.current_path || '';

            tableEl.style.display = 'block';
            renderDocumentsTable(cachedData.documents);
            renderBreadcrumb(AppState.currentPath);

            const message = `${cachedData.folders_count} pastas, ${cachedData.files_count} arquivos (cache)`;
            showToast(message, 'success');
            return;
        }
    }

    // Mostra loading
    loadingEl.style.display = 'block';
    loadingEl.innerHTML = '<div class="spinner"></div><p>Carregando todos os documentos e bibliotecas do SharePoint...<br><small>Isso pode levar alguns segundos se houver muitos arquivos</small></p>';
    tableEl.style.display = 'none';
    noDocsEl.style.display = 'none';

    try {
        // Constrói URL com parâmetro de pasta se fornecido
        let url = '/api/sharepoint/documents';
        const params = [];

        if (folderPath) {
            params.push(`folder_path=${encodeURIComponent(folderPath)}`);
        }

        if (forceRefresh) {
            params.push('force_refresh=true');
        }

        if (params.length > 0) {
            url += '?' + params.join('&');
        }

        const response = await fetch(url, {
            credentials: 'include'
        });

        if (!response.ok) {
            throw new Error('Erro ao carregar documentos');
        }

        const data = await response.json();
        AppState.documents = data.documents;
        AppState.currentPath = data.current_path || '';

        // Salva no cache local
        CacheManager.set(cacheKey, data);

        // Restaura loading HTML e esconde
        loadingEl.innerHTML = '<div class="spinner"></div><p>Carregando documentos...</p>';
        loadingEl.style.display = 'none';

        if (data.documents.length === 0) {
            noDocsEl.style.display = 'block';
        } else {
            tableEl.style.display = 'block';
            renderDocumentsTable(data.documents);
            renderBreadcrumb(AppState.currentPath);

            const cacheIndicator = data.from_cache ? ' (cache servidor)' : '';
            const message = `${data.folders_count} pastas, ${data.files_count} arquivos${cacheIndicator}`;
            showToast(message, 'success');
        }
    } catch (error) {
        console.error('Erro ao carregar documentos:', error);
        loadingEl.innerHTML = '<div class="spinner"></div><p>Carregando documentos...</p>';
        loadingEl.style.display = 'none';
        showToast('Erro ao carregar documentos do SharePoint', 'error');
    }
}

function renderDocumentsTable(documents) {
    const tbody = document.getElementById('documentsTableBody');
    tbody.innerHTML = '';

    documents.forEach(doc => {
        const row = document.createElement('tr');
        row.setAttribute('data-doc-name', doc.name.toLowerCase());
        row.setAttribute('data-is-folder', doc.isFolder ? 'true' : 'false');

        if (doc.isFolder) {
            // Renderiza linha de pasta
            row.classList.add('folder-row');
            const folderPath = doc.folderPath || doc.name;
            const isDrive = doc.isDrive || false;
            const folderIcon = isDrive ? '📚' : '📁';

            row.innerHTML = `
                <td>
                    <span class="folder-icon">${folderIcon}</span>
                </td>
                <td>
                    <strong class="folder-name" data-folder-path="${folderPath}" style="cursor: pointer; color: var(--primary-color);">
                        ${doc.name}
                    </strong>
                    ${doc.driveName && !isDrive ? `<br><small class="text-muted">📚 ${doc.driveName}</small>` : ''}
                </td>
                <td>
                    <span class="file-type-badge badge-folder">
                        ${isDrive ? 'Biblioteca' : 'Pasta'}
                    </span>
                </td>
                <td><span class="text-muted">${doc.childCount} itens</span></td>
                <td>${formatDate(doc.lastModified)}</td>
                <td>
                    <a href="${doc.webUrl}" target="_blank" class="btn btn-sm btn-secondary">
                        Abrir
                    </a>
                </td>
            `;
        } else {
            // Renderiza linha de arquivo
            row.innerHTML = `
                <td>
                    <input type="checkbox" class="doc-checkbox" data-doc-id="${doc.id}"
                           data-doc-name="${doc.name}"
                           data-doc-type="${doc.type}"
                           data-drive-id="${doc.driveId}">
                </td>
                <td>
                    <strong>${doc.name}</strong>
                    ${doc.driveName ? `<br><small class="text-muted">📚 ${doc.driveName}</small>` : ''}
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
        }

        tbody.appendChild(row);
    });

    // Adiciona event listeners aos checkboxes
    document.querySelectorAll('.doc-checkbox').forEach(checkbox => {
        checkbox.addEventListener('change', handleDocumentSelection);
    });

    // Adiciona event listeners para clique em pastas
    document.querySelectorAll('.folder-name').forEach(folderEl => {
        folderEl.addEventListener('click', (e) => {
            const folderPath = e.target.getAttribute('data-folder-path');
            navigateToFolder(folderPath);
        });
    });

    // Atualiza contador de pesquisa
    updateSearchResults();
}

function navigateToFolder(folderPath) {
    // Navega para a pasta usando o caminho fornecido pelo backend
    loadDocuments(folderPath);
}

function renderBreadcrumb(currentPath) {
    const breadcrumbEl = document.getElementById('breadcrumb');
    const breadcrumbContainer = document.getElementById('breadcrumbContainer');

    if (!breadcrumbEl) return;

    // Mostra o breadcrumb
    if (breadcrumbContainer) {
        breadcrumbContainer.style.display = 'block';
    }

    breadcrumbEl.innerHTML = '';

    // Botão Home
    const homeLink = document.createElement('span');
    homeLink.className = 'breadcrumb-item';
    homeLink.innerHTML = '<span class="breadcrumb-link">🏠 Raiz</span>';
    homeLink.style.cursor = 'pointer';
    homeLink.addEventListener('click', () => loadDocuments(null));
    breadcrumbEl.appendChild(homeLink);

    // Se não há caminho, para por aqui
    if (!currentPath) {
        homeLink.classList.add('breadcrumb-active');
        return;
    }

    // Parse do caminho (formato: drive:xxx/path/to/folder)
    let pathParts = [];
    let driveId = null;

    if (currentPath.startsWith('drive:')) {
        const parts = currentPath.substring(6).split('/');
        driveId = parts[0];
        pathParts = parts.slice(1);
    } else {
        // Fallback para formato antigo
        pathParts = currentPath.split('/').filter(p => p);
    }

    // Se temos um drive, adiciona link para a raiz do drive
    if (driveId) {
        const separator = document.createElement('span');
        separator.className = 'breadcrumb-separator';
        separator.textContent = '/';
        breadcrumbEl.appendChild(separator);

        const driveLink = document.createElement('span');
        driveLink.className = 'breadcrumb-item';

        // Se não há subpastas, este é o item ativo
        if (pathParts.length === 0) {
            driveLink.innerHTML = `<span class="breadcrumb-active">📚 Biblioteca</span>`;
        } else {
            driveLink.innerHTML = `<span class="breadcrumb-link">📚 Biblioteca</span>`;
            driveLink.style.cursor = 'pointer';
            driveLink.addEventListener('click', () => loadDocuments(`drive:${driveId}`));
        }

        breadcrumbEl.appendChild(driveLink);
    }

    // Adiciona as subpastas
    pathParts.forEach((part, index) => {
        // Adiciona separador
        const separator = document.createElement('span');
        separator.className = 'breadcrumb-separator';
        separator.textContent = '/';
        breadcrumbEl.appendChild(separator);

        // Adiciona link da pasta
        const link = document.createElement('span');
        link.className = 'breadcrumb-item';

        // Constrói o caminho até esta pasta
        const pathToHere = `drive:${driveId}/${pathParts.slice(0, index + 1).join('/')}`;

        if (index === pathParts.length - 1) {
            // Última pasta (atual) - não é clicável
            link.innerHTML = `<span class="breadcrumb-active">${part}</span>`;
        } else {
            // Pasta intermediária - é clicável
            link.innerHTML = `<span class="breadcrumb-link">${part}</span>`;
            link.style.cursor = 'pointer';
            link.addEventListener('click', () => loadDocuments(pathToHere));
        }

        breadcrumbEl.appendChild(link);
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

    // Chave de cache para esta combinação de documentos
    const docIds = AppState.selectedDocuments.map(d => d.id).sort().join('_');
    const cacheKey = `summary_${docIds}`;

    // Tenta obter do cache local
    const cachedSummary = CacheManager.get(cacheKey);
    if (cachedSummary) {
        summariesSection.style.display = 'block';
        summariesLoading.style.display = 'none';
        summariesContent.innerHTML = '';

        renderSummaries(cachedSummary.summaries, cachedSummary.consolidated_summary);
        showToast('Resumos carregados do cache!', 'success');

        // Scroll suave para seção de resumos
        summariesSection.scrollIntoView({ behavior: 'smooth', block: 'start' });
        return;
    }

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

        // Salva no cache local (30 minutos)
        CacheManager.set(cacheKey, {
            summaries: data.summaries,
            consolidated_summary: data.consolidated_summary
        });

        renderSummaries(data.summaries, data.consolidated_summary);

        const cacheInfo = data.cache_hits > 0 ? ` (${data.cache_hits} do cache)` : '';
        showToast(`Resumos gerados com sucesso!${cacheInfo}`, 'success');

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
// Pesquisa de Documentos
// ============================================

function handleSearch(e) {
    const searchTerm = e.target.value.toLowerCase().trim();
    const clearBtn = document.getElementById('clearSearchBtn');
    const rows = document.querySelectorAll('#documentsTableBody tr');

    // Mostra/esconde botão de limpar
    clearBtn.style.display = searchTerm ? 'block' : 'none';

    // Filtra documentos
    rows.forEach(row => {
        const docName = row.getAttribute('data-doc-name');
        if (docName.includes(searchTerm)) {
            row.style.display = '';
        } else {
            row.style.display = 'none';
            // Desmarca checkbox se estiver selecionado
            const checkbox = row.querySelector('.doc-checkbox');
            if (checkbox && checkbox.checked) {
                checkbox.checked = false;
            }
        }
    });

    // Atualiza contador de resultados
    updateSearchResults();

    // Atualiza seleção
    updateSelectedDocuments();
}

function clearSearch() {
    const searchInput = document.getElementById('searchInput');
    searchInput.value = '';
    searchInput.dispatchEvent(new Event('input'));
}

function updateSearchResults() {
    const searchResults = document.getElementById('searchResults');
    const searchTerm = document.getElementById('searchInput').value.trim();
    const allRows = document.querySelectorAll('#documentsTableBody tr');
    const visibleRows = document.querySelectorAll('#documentsTableBody tr[style=""]');
    const totalCount = allRows.length;
    const visibleCount = visibleRows.length;

    if (searchTerm && visibleCount < totalCount) {
        searchResults.textContent = `Mostrando ${visibleCount} de ${totalCount} documentos`;
        searchResults.style.display = 'block';
    } else {
        searchResults.textContent = '';
        searchResults.style.display = 'none';
    }
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
        'text': 'Texto',
        'folder': 'Pasta'
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
