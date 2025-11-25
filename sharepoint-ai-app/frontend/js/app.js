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
    currentWebUrl: '',
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

    // Navigation
    document.getElementById('backBtn').addEventListener('click', handleBack);
    document.getElementById('copyPathBtn').addEventListener('click', handleCopyPath);
    document.getElementById('openInSharePointBtn').addEventListener('click', handleOpenInSharePoint);
    document.getElementById('pathInput').addEventListener('click', handleCopyPath);

    // Search
    document.getElementById('searchInput').addEventListener('input', handleSearch);
    document.getElementById('searchInput').addEventListener('keydown', handleSearch);

    // Chat
    const chatSend = document.getElementById('chatSend');
    const chatInput = document.getElementById('chatInput');

    chatSend.addEventListener('click', () => sendChatMessage());

    // Enter to send (Shift+Enter for new line)
    chatInput.addEventListener('keydown', (e) => {
        if (e.key === 'Enter' && !e.shiftKey) {
            e.preventDefault();
            sendChatMessage();
        }
    });

    // Auto-resize textarea
    chatInput.addEventListener('input', () => {
        chatInput.style.height = 'auto';
        chatInput.style.height = Math.min(chatInput.scrollHeight, 120) + 'px';
    });
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
    document.getElementById('loginSection').style.display = 'flex';
    document.getElementById('mainApp').style.display = 'none';
}

function showDocumentsSection() {
    document.getElementById('loginSection').style.display = 'none';
    document.getElementById('mainApp').style.display = 'flex';

    // Atualiza informações do usuário
    document.getElementById('userName').textContent = AppState.user.name;

    // Cria avatar com iniciais
    const initials = AppState.user.name.split(' ').map(n => n[0]).join('').substring(0, 2).toUpperCase();
    document.getElementById('userAvatar').textContent = initials;
}

// ============================================
// Documentos
// ============================================

async function loadDocuments(folderPath = null, forceRefresh = false, skipHistory = false) {
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
            updatePathBar(AppState.currentPath, cachedData.folder_name, cachedData.web_url);

            const message = `${cachedData.folders_count} pastas, ${cachedData.files_count} arquivos (cache)`;
            showToast(message, 'success');
            return;
        }
    }

    // Mostra loading
    loadingEl.style.display = 'flex';
    loadingEl.innerHTML = `
        <div class="spinner mb-4"></div>
        <p class="text-gray-600 dark:text-gray-400 text-center">
            Carregando todos os documentos e bibliotecas do SharePoint...<br>
            <small class="text-gray-500 dark:text-gray-500">Isso pode levar alguns segundos se houver muitos arquivos</small>
        </p>
    `;
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
        loadingEl.innerHTML = `
            <div class="spinner mb-4"></div>
            <p class="text-gray-600 dark:text-gray-400">Carregando documentos...</p>
        `;
        loadingEl.style.display = 'none';

        if (data.documents.length === 0) {
            noDocsEl.style.display = 'block';
        } else {
            tableEl.style.display = 'block';
            renderDocumentsTable(data.documents);
            renderBreadcrumb(AppState.currentPath);
            updatePathBar(AppState.currentPath, data.folder_name, data.web_url);

            const cacheIndicator = data.from_cache ? ' (cache servidor)' : '';
            const message = `${data.folders_count} pastas, ${data.files_count} arquivos${cacheIndicator}`;
            showToast(message, 'success');
        }
    } catch (error) {
        console.error('Erro ao carregar documentos:', error);
        loadingEl.innerHTML = `
            <div class="spinner mb-4"></div>
            <p class="text-gray-600 dark:text-gray-400">Carregando documentos...</p>
        `;
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

        row.className = 'border-b border-neutral-200 dark:border-neutral-700 hover:bg-neutral-100 dark:hover:bg-neutral-800/50';

        if (doc.isFolder) {
            // Renderiza linha de pasta
            const folderPath = doc.folderPath || doc.name;
            const isDrive = doc.isDrive || false;
            const folderIcon = isDrive ? 'folder_special' : 'folder';
            const iconColor = isDrive ? 'text-yellow-500' : 'text-blue-500';

            row.innerHTML = `
                <td class="p-3">
                    <span class="material-symbols-outlined ${iconColor} text-2xl">${folderIcon}</span>
                </td>
                <td class="p-3">
                    <strong class="folder-name cursor-pointer text-primary-600 dark:text-primary-400 hover:underline" data-folder-path="${folderPath}">
                        ${doc.name}
                    </strong>
                    ${doc.driveName && !isDrive ? `<br><small class="text-gray-500 dark:text-gray-400 flex items-center gap-1 mt-1"><span class="material-symbols-outlined text-xs">folder_special</span>${doc.driveName}</small>` : ''}
                </td>
                <td class="p-3 text-gray-600 dark:text-gray-400">
                    ${doc.childCount} itens
                </td>
                <td class="p-3 text-gray-600 dark:text-gray-400">
                    ${formatDate(doc.lastModified)}
                </td>
                <td class="p-3 text-right">
                    <a href="${doc.webUrl}" target="_blank" class="px-3 py-1 text-sm bg-neutral-200 dark:bg-neutral-700 rounded hover:bg-neutral-300 dark:hover:bg-neutral-600">
                        Abrir
                    </a>
                </td>
            `;
        } else {
            // Renderiza linha de arquivo
            const fileIcon = getFileMaterialIcon(doc.type);
            const iconColor = getFileIconColor(doc.type);

            row.innerHTML = `
                <td class="p-3">
                    <input type="checkbox" class="doc-checkbox rounded border-gray-300 dark:border-gray-600 text-primary-600 focus:ring-primary-500"
                           data-doc-id="${doc.id}"
                           data-doc-name="${doc.name}"
                           data-doc-type="${doc.type}"
                           data-drive-id="${doc.driveId}">
                </td>
                <td class="p-3">
                    <div class="flex items-center gap-3">
                        <span class="material-symbols-outlined ${iconColor} text-2xl">${fileIcon}</span>
                        <div>
                            <strong class="text-gray-900 dark:text-gray-100">${doc.name}</strong>
                            ${doc.driveName ? `<br><small class="text-gray-500 dark:text-gray-400 flex items-center gap-1 mt-1"><span class="material-symbols-outlined text-xs">folder_special</span>${doc.driveName}</small>` : ''}
                        </div>
                    </div>
                </td>
                <td class="p-3 text-gray-600 dark:text-gray-400">${formatDate(doc.lastModified)}</td>
                <td class="p-3 text-gray-600 dark:text-gray-400">${formatFileSize(doc.size)}</td>
                <td class="p-3 text-right">
                    <a href="${doc.webUrl}" target="_blank" class="px-3 py-1 text-sm bg-neutral-200 dark:bg-neutral-700 rounded hover:bg-neutral-300 dark:hover:bg-neutral-600">
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
}

function navigateToFolder(folderPath) {
    // Adiciona ao histórico de navegação antes de navegar
    if (AppState.currentPath !== folderPath) {
        AppState.navigationHistory.push(AppState.currentPath);
        updateBackButton();
    }

    // Navega para a pasta usando o caminho fornecido pelo backend
    loadDocuments(folderPath);
}

function handleBack() {
    if (AppState.navigationHistory.length > 0) {
        const previousPath = AppState.navigationHistory.pop();
        updateBackButton();

        // Carrega documentos sem adicionar ao histórico
        loadDocuments(previousPath, false, true);
    }
}

function updateBackButton() {
    const backBtn = document.getElementById('backBtn');
    backBtn.disabled = AppState.navigationHistory.length === 0;
}

function updatePathBar(currentPath, folderName = null, webUrl = null) {
    const pathInput = document.getElementById('pathInput');

    // Salva a URL do SharePoint
    if (webUrl) {
        AppState.currentWebUrl = webUrl;
    }

    // Atualiza o campo de caminho
    if (!currentPath) {
        pathInput.value = '/';
    } else {
        // Extrai o nome legível do caminho
        if (currentPath.startsWith('drive:')) {
            const parts = currentPath.substring(6).split('/');
            const pathParts = parts.slice(1); // Remove o drive ID

            if (pathParts.length === 0) {
                pathInput.value = folderName || '/Biblioteca';
            } else {
                pathInput.value = '/' + pathParts.join('/');
            }
        } else {
            pathInput.value = currentPath;
        }
    }
}

async function handleCopyPath() {
    const pathInput = document.getElementById('pathInput');
    const path = pathInput.value;

    try {
        await navigator.clipboard.writeText(path);

        // Feedback visual
        pathInput.select();
        showToast('Caminho copiado para a área de transferência', 'success');

        // Remove seleção após um tempo
        setTimeout(() => {
            window.getSelection().removeAllRanges();
        }, 500);
    } catch (err) {
        console.error('Erro ao copiar:', err);
        showToast('Erro ao copiar caminho', 'error');
    }
}

function handleOpenInSharePoint() {
    if (AppState.currentWebUrl) {
        window.open(AppState.currentWebUrl, '_blank');
    } else {
        // URL base do SharePoint
        const baseUrl = 'https://fiofortei9automacaogroup.sharepoint.com/sites/Documentos';
        window.open(baseUrl, '_blank');
    }
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

    // Limpa breadcrumb e adiciona home
    breadcrumbEl.innerHTML = '';

    // Botão Home
    const homeSpan = document.createElement('span');
    homeSpan.className = 'cursor-pointer hover:text-primary-600';
    homeSpan.innerHTML = '<span class="material-symbols-outlined text-xl">home</span>';
    homeSpan.addEventListener('click', () => loadDocuments(null));
    breadcrumbEl.appendChild(homeSpan);

    // Se não há caminho, para por aqui
    if (!currentPath) {
        const rootText = document.createElement('span');
        rootText.className = 'text-primary-600 font-semibold';
        rootText.innerHTML = '<span>/</span><span>Raiz</span>';
        breadcrumbEl.appendChild(rootText);
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
        breadcrumbEl.innerHTML += ' <span>/</span> ';

        const driveLink = document.createElement('span');

        // Se não há subpastas, este é o item ativo
        if (pathParts.length === 0) {
            driveLink.className = 'text-primary-600 font-semibold';
            driveLink.textContent = 'Biblioteca';
        } else {
            driveLink.className = 'cursor-pointer hover:text-primary-600';
            driveLink.textContent = 'Biblioteca';
            driveLink.addEventListener('click', () => loadDocuments(`drive:${driveId}`));
        }

        breadcrumbEl.appendChild(driveLink);
    }

    // Adiciona as subpastas
    pathParts.forEach((part, index) => {
        breadcrumbEl.innerHTML += ' <span>/</span> ';

        // Adiciona link da pasta
        const link = document.createElement('span');

        // Constrói o caminho até esta pasta
        const pathToHere = `drive:${driveId}/${pathParts.slice(0, index + 1).join('/')}`;

        if (index === pathParts.length - 1) {
            // Última pasta (atual) - não é clicável
            link.className = 'text-primary-600 font-semibold';
            link.textContent = part;
        } else {
            // Pasta intermediária - é clicável
            link.className = 'cursor-pointer hover:text-primary-600';
            link.textContent = part;
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
    const count = AppState.selectedDocuments.length;
    document.getElementById('selectedCount').textContent = `${count} selecionado${count !== 1 ? 's' : ''}`;

    const summarizeBtn = document.getElementById('summarizeBtn');
    summarizeBtn.disabled = count === 0;

    const summarizeBtnText = document.getElementById('summarizeBtnText');
    if (summarizeBtnText) {
        summarizeBtnText.textContent = count > 0 ? `Resumir ${count} Documento${count !== 1 ? 's' : ''}` : 'Resumir Documentos';
    }
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

// Variável para controlar timeout de busca
let searchTimeout = null;
let isGlobalSearchActive = false;

function handleSearch(e) {
    const searchTerm = e.target.value.trim();

    // Se não há busca, volta para filtro local
    if (!searchTerm) {
        isGlobalSearchActive = false;
        const rows = document.querySelectorAll('#documentsTableBody tr');
        rows.forEach(row => row.style.display = '');
        updateSelectedDocuments();
        return;
    }

    // Limpa timeout anterior
    if (searchTimeout) {
        clearTimeout(searchTimeout);
    }

    // Se Enter foi pressionado, busca imediatamente
    if (e.key === 'Enter') {
        performGlobalSearch(searchTerm);
        return;
    }

    // Caso contrário, faz busca local primeiro (instantânea)
    performLocalSearch(searchTerm);

    // Depois de 1 segundo sem digitar, dispara busca global
    searchTimeout = setTimeout(() => {
        performGlobalSearch(searchTerm);
    }, 1000);
}

function performLocalSearch(searchTerm) {
    const searchTermLower = searchTerm.toLowerCase();
    const rows = document.querySelectorAll('#documentsTableBody tr');

    let visibleCount = 0;

    // Filtra documentos localmente
    rows.forEach(row => {
        const docName = row.getAttribute('data-doc-name');
        if (docName.includes(searchTermLower)) {
            row.style.display = '';
            visibleCount++;
        } else {
            row.style.display = 'none';
            // Desmarca checkbox se estiver selecionado
            const checkbox = row.querySelector('.doc-checkbox');
            if (checkbox && checkbox.checked) {
                checkbox.checked = false;
            }
        }
    });

    // Atualiza contador de selecionados
    updateSelectedDocuments();

    // Mostra toast com resultado da busca local
    if (visibleCount < rows.length) {
        showToast(`Busca local: ${visibleCount} de ${rows.length} documentos`, 'info');
    }
}

async function performGlobalSearch(searchTerm) {
    try {
        isGlobalSearchActive = true;
        showToast('🔍 Buscando em todo o SharePoint...', 'info');

        const response = await fetch(`/api/sharepoint/search?q=${encodeURIComponent(searchTerm)}&limit=200`);
        const data = await response.json();

        if (data.success) {
            // Atualiza AppState com resultados
            AppState.documents = data.documents;
            AppState.selectedDocuments = [];

            // Renderiza documentos encontrados
            renderDocumentsTable(data.documents);

            // Atualiza mensagem de resultados
            const cacheText = data.from_cache ? ' (cache)' : '';
            showToast(`Busca global: ${data.count} documentos encontrados em todo o SharePoint${cacheText}`, 'success');
        } else {
            showToast(`Erro na busca: ${data.error}`, 'error');
        }

    } catch (error) {
        console.error('Erro na busca global:', error);
        showToast('Erro ao buscar documentos', 'error');
    }
}

function clearSearch() {
    const searchInput = document.getElementById('searchInput');
    searchInput.value = '';

    // Se estava em busca global, recarrega a pasta atual
    if (isGlobalSearchActive) {
        isGlobalSearchActive = false;
        loadDocuments(AppState.currentPath);
    } else {
        searchInput.dispatchEvent(new Event('input'));
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

function getFileMaterialIcon(type) {
    const icons = {
        'pdf': 'picture_as_pdf',
        'word': 'description',
        'excel': 'table_chart',
        'powerpoint': 'slideshow',
        'text': 'description',
        'image': 'image',
        'video': 'video_file',
        'audio': 'audio_file'
    };
    return icons[type] || 'description';
}

function getFileIconColor(type) {
    const colors = {
        'pdf': 'text-red-500',
        'word': 'text-blue-500',
        'excel': 'text-green-500',
        'powerpoint': 'text-orange-500',
        'text': 'text-gray-500',
        'image': 'text-purple-500',
        'video': 'text-pink-500',
        'audio': 'text-indigo-500'
    };
    return colors[type] || 'text-gray-500';
}

function showToast(message, type = 'info') {
    const container = document.getElementById('toastContainer');
    const toast = document.createElement('div');

    const configs = {
        'success': {
            bg: 'bg-green-500',
            icon: 'check_circle',
            text: 'text-white'
        },
        'error': {
            bg: 'bg-red-500',
            icon: 'error',
            text: 'text-white'
        },
        'info': {
            bg: 'bg-blue-500',
            icon: 'info',
            text: 'text-white'
        }
    };

    const config = configs[type] || configs['info'];

    toast.className = `flex items-center gap-3 px-4 py-3 rounded-lg shadow-lg ${config.bg} ${config.text} transform transition-all duration-300 opacity-0 translate-x-full`;

    toast.innerHTML = `
        <span class="material-symbols-outlined">${config.icon}</span>
        <span class="flex-1">${message}</span>
    `;

    container.appendChild(toast);

    // Animate in
    setTimeout(() => {
        toast.classList.remove('opacity-0', 'translate-x-full');
    }, 10);

    // Remove toast após 5 segundos
    setTimeout(() => {
        toast.classList.add('opacity-0', 'translate-x-full');
        setTimeout(() => toast.remove(), 300);
    }, 5000);
}

// ============================================
// Chat AI Widget
// ============================================

const ChatState = {
    messages: [],
    isProcessing: false
};

async function sendChatMessage() {
    const chatInput = document.getElementById('chatInput');
    const message = chatInput.value.trim();

    if (!message || ChatState.isProcessing) return;

    // Verifica autenticação
    if (!AppState.authenticated) {
        showToast('Faça login primeiro para usar o chat', 'error');
        return;
    }

    console.log('%c╔════════════════════════════════════════════════════════════════╗', 'color: #0078d4; font-weight: bold');
    console.log('%c║ 🤖 SOFIA - Iniciando processamento                            ║', 'color: #0078d4; font-weight: bold');
    console.log('%c╚════════════════════════════════════════════════════════════════╝', 'color: #0078d4; font-weight: bold');
    console.log('📝 Pergunta do usuário:', message);
    console.log('⏰ Timestamp:', new Date().toLocaleTimeString());

    // Limpa input
    chatInput.value = '';
    chatInput.style.height = 'auto';

    // Adiciona mensagem do usuário
    addChatMessage('user', message);

    // Mostra loading
    ChatState.isProcessing = true;
    showChatLoading(true);

    // Cria mensagem de "pensando" com etapas
    const thinkingMessage = addThinkingMessage();

    try {
        // Etapa 1: Iniciando
        console.log('%c⏳ ETAPA 1: Processando pergunta', 'color: #0078d4; font-weight: bold');
        updateThinkingStep(thinkingMessage, 0, 'active');
        await sleep(300);

        // Etapa 2: Buscando documentos
        console.log('%c✓ ETAPA 1: Concluída', 'color: #107c10; font-weight: bold');
        console.log('%c🔍 ETAPA 2: Buscando documentos no SharePoint', 'color: #0078d4; font-weight: bold');
        updateThinkingStep(thinkingMessage, 0, 'completed');
        updateThinkingStep(thinkingMessage, 1, 'active');

        const requestPayload = {
            message: message,
            conversation_history: ChatState.messages.slice(-5)
        };
        console.log('📤 Enviando requisição para API:', requestPayload);

        const response = await fetch('/api/ai/chat', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            credentials: 'include',
            body: JSON.stringify(requestPayload)
        });

        console.log('📥 Resposta recebida - Status:', response.status);

        const data = await response.json();
        console.log('📦 Dados recebidos:', data);

        if (data.success) {
            // Etapa 3: Analisando documentos
            console.log('%c✓ ETAPA 2: Concluída', 'color: #107c10; font-weight: bold');
            console.log('%c📄 ETAPA 3: Analisando documentos', 'color: #0078d4; font-weight: bold');
            console.log(`   📊 Documentos encontrados: ${data.documents_found}`);
            console.log(`   🔍 Documentos analisados: ${data.documents_analyzed}`);

            if (data.sources && data.sources.length > 0) {
                console.log('   📄 Fontes utilizadas:');
                data.sources.forEach((source, index) => {
                    console.log(`      ${index + 1}. ${source.name} (${source.type})`);
                    console.log(`         URL: ${source.webUrl}`);
                });
            }

            updateThinkingStep(thinkingMessage, 1, 'completed');
            updateThinkingStep(thinkingMessage, 2, 'active');
            updateThinkingStepText(thinkingMessage, 2, `Analisando ${data.documents_analyzed} documentos encontrados`);
            await sleep(500);

            // Etapa 4: Gerando resposta
            console.log('%c✓ ETAPA 3: Concluída', 'color: #107c10; font-weight: bold');
            console.log('%c✨ ETAPA 4: Gerando resposta com Claude AI', 'color: #0078d4; font-weight: bold');
            console.log('   💬 Resposta gerada:', data.response.substring(0, 100) + '...');

            updateThinkingStep(thinkingMessage, 2, 'completed');
            updateThinkingStep(thinkingMessage, 3, 'active');
            await sleep(500);

            // Todas etapas completas
            console.log('%c✓ ETAPA 4: Concluída', 'color: #107c10; font-weight: bold');
            console.log('%c✅ PROCESSAMENTO COMPLETO!', 'color: #107c10; font-weight: bold; font-size: 14px');
            console.log('%c════════════════════════════════════════════════════════════════', 'color: #107c10; font-weight: bold');

            updateThinkingStep(thinkingMessage, 3, 'completed');
            await sleep(300);

            // Remove mensagem de pensamento
            thinkingMessage.remove();

            // Adiciona resposta do bot com fontes
            addChatMessage('bot', data.response, data.sources || [], data.documents_analyzed, data.documents_found);
        } else {
            console.error('%c❌ ERRO NA API:', 'color: #d13438; font-weight: bold', data.error);
            thinkingMessage.remove();
            addChatMessage('bot', `❌ Erro: ${data.error}`);
        }

    } catch (error) {
        console.error('%c❌ ERRO CRÍTICO:', 'color: #d13438; font-weight: bold');
        console.error('Detalhes do erro:', error);
        console.error('Stack trace:', error.stack);
        thinkingMessage.remove();
        addChatMessage('bot', '❌ Erro ao processar sua mensagem. Tente novamente.');
    } finally {
        ChatState.isProcessing = false;
        showChatLoading(false);
        console.log('%c════════════════════════════════════════════════════════════════', 'color: #666');
        console.log(' ');
    }
}

function sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}

function addThinkingMessage() {
    const chatMessages = document.getElementById('chatMessages');
    const messageDiv = document.createElement('div');
    messageDiv.className = 'flex gap-3';
    messageDiv.id = 'thinking-message';

    messageDiv.innerHTML = `
        <div class="flex-shrink-0">
            <div class="w-8 h-8 rounded-full bg-primary-100 dark:bg-primary-900 flex items-center justify-center">
                <span class="material-symbols-outlined text-primary-600 dark:text-primary-400 text-lg">smart_toy</span>
            </div>
        </div>
        <div class="flex-1 bg-neutral-100 dark:bg-neutral-800 rounded-lg p-3">
            <p class="text-sm font-semibold mb-2">Sofia está pensando...</p>
            <div class="space-y-2 text-sm text-gray-600 dark:text-gray-400">
                <div class="thinking-step flex items-center gap-2" data-step="0">
                    <span class="material-symbols-outlined text-lg">hourglass_empty</span>
                    <span class="step-text">Processando sua pergunta</span>
                </div>
                <div class="thinking-step flex items-center gap-2" data-step="1">
                    <span class="material-symbols-outlined text-lg">search</span>
                    <span class="step-text">Buscando documentos relevantes no SharePoint</span>
                </div>
                <div class="thinking-step flex items-center gap-2" data-step="2">
                    <span class="material-symbols-outlined text-lg">description</span>
                    <span class="step-text">Analisando conteúdo dos documentos</span>
                </div>
                <div class="thinking-step flex items-center gap-2" data-step="3">
                    <span class="material-symbols-outlined text-lg">auto_awesome</span>
                    <span class="step-text">Gerando resposta inteligente</span>
                </div>
            </div>
        </div>
    `;

    chatMessages.appendChild(messageDiv);
    chatMessages.scrollTop = chatMessages.scrollHeight;

    return messageDiv;
}

function updateThinkingStep(messageDiv, stepIndex, status) {
    const step = messageDiv.querySelector(`[data-step="${stepIndex}"]`);
    if (!step) return;

    step.classList.remove('active', 'completed');
    if (status) {
        step.classList.add(status);
    }
}

function updateThinkingStepText(messageDiv, stepIndex, newText) {
    const step = messageDiv.querySelector(`[data-step="${stepIndex}"]`);
    if (!step) return;

    const textSpan = step.querySelector('.step-text');
    if (textSpan) {
        textSpan.textContent = newText;
    }
}

function addChatMessage(sender, content, sources = [], docsAnalyzed = 0, docsFound = 0) {
    const chatMessages = document.getElementById('chatMessages');
    const messageDiv = document.createElement('div');
    messageDiv.className = 'flex gap-3';

    const isUser = sender === 'user';
    const avatarBg = isUser ? 'bg-gray-200 dark:bg-gray-700' : 'bg-primary-100 dark:bg-primary-900';
    const avatarColor = isUser ? 'text-gray-600 dark:text-gray-300' : 'text-primary-600 dark:text-primary-400';
    const avatarIcon = isUser ? 'person' : 'smart_toy';
    const messageBg = isUser ? 'bg-primary-50 dark:bg-primary-900/20' : 'bg-neutral-100 dark:bg-neutral-800';

    // Adiciona estatísticas se houver
    let statsHTML = '';
    if (!isUser && docsAnalyzed > 0) {
        statsHTML = `
            <div class="text-xs text-gray-500 dark:text-gray-400 mt-2 flex items-center gap-2">
                <span class="material-symbols-outlined text-sm">analytics</span>
                <span>${docsFound} documentos encontrados • ${docsAnalyzed} analisados</span>
            </div>
        `;
    }

    let sourcesHTML = '';
    if (sources && sources.length > 0) {
        sourcesHTML = `
            <div class="mt-3 pt-3 border-t border-gray-200 dark:border-gray-700">
                <div class="text-xs font-semibold text-gray-600 dark:text-gray-400 mb-2 flex items-center gap-1">
                    <span class="material-symbols-outlined text-sm">description</span>
                    <span>Fontes consultadas:</span>
                </div>
                <div class="space-y-1">
                    ${sources.map(source => {
                        const icon = getFileMaterialIcon(source.type);
                        const iconColor = getFileIconColor(source.type);
                        return `
                            <a href="${source.webUrl}" target="_blank" class="flex items-center gap-2 text-xs text-primary-600 dark:text-primary-400 hover:underline">
                                <span class="material-symbols-outlined ${iconColor} text-sm">${icon}</span>
                                <span>${source.name}</span>
                            </a>
                        `;
                    }).join('')}
                </div>
            </div>
        `;
    }

    messageDiv.innerHTML = `
        <div class="flex-shrink-0">
            <div class="w-8 h-8 rounded-full ${avatarBg} flex items-center justify-center">
                <span class="material-symbols-outlined ${avatarColor} text-lg">${avatarIcon}</span>
            </div>
        </div>
        <div class="flex-1 ${messageBg} rounded-lg p-3">
            <div class="text-sm">${formatChatMessage(content)}</div>
            ${statsHTML}
            ${sourcesHTML}
        </div>
    `;

    chatMessages.appendChild(messageDiv);

    // Scroll to bottom
    chatMessages.scrollTop = chatMessages.scrollHeight;

    // Salva mensagem no histórico
    ChatState.messages.push({
        sender: sender,
        content: content,
        timestamp: Date.now()
    });
}

function formatChatMessage(text) {
    // Converte markdown básico para HTML
    let formatted = text
        .replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>')
        .replace(/\n/g, '<br>');

    // Detecta listas
    if (formatted.includes('- ')) {
        const lines = formatted.split('<br>');
        let inList = false;
        let result = [];

        for (let line of lines) {
            if (line.trim().startsWith('- ')) {
                if (!inList) {
                    result.push('<ul>');
                    inList = true;
                }
                result.push(`<li>${line.trim().substring(2)}</li>`);
            } else {
                if (inList) {
                    result.push('</ul>');
                    inList = false;
                }
                result.push(line);
            }
        }

        if (inList) {
            result.push('</ul>');
        }

        formatted = result.join('');
    }

    return `<p>${formatted}</p>`;
}

function showChatLoading(show) {
    const chatLoading = document.getElementById('chatLoading');
    const chatSend = document.getElementById('chatSend');

    chatLoading.style.display = show ? 'flex' : 'none';
    chatSend.disabled = show;
}

function getFileIcon(type) {
    const icons = {
        'pdf': '📕',
        'word': '📘',
        'excel': '📊',
        'powerpoint': '📽️',
        'text': '📄',
        'folder': '📁',
        'image': '🖼️',
        'video': '🎥',
        'audio': '🎵'
    };
    return icons[type] || '📄';
}

// Chat widget já está integrado no template, não precisa inicializar
