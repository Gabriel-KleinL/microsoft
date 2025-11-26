/**
 * Recentes Page - JavaScript
 */

// Importa o RecentDocumentsManager do app.js (assumindo que está disponível globalmente)

document.addEventListener('DOMContentLoaded', () => {
    checkAuthStatus();
    loadRecentDocuments();
});

async function checkAuthStatus() {
    try {
        const response = await fetch('/api/auth/status', {
            credentials: 'include'
        });
        const data = await response.json();

        if (data.authenticated) {
            const userName = data.user_name;
            const userEmail = data.user_email;

            document.getElementById('userName').textContent = userName;

            // Cria avatar com iniciais
            const initials = userName.split(' ').map(n => n[0]).join('').substring(0, 2).toUpperCase();
            document.getElementById('userAvatar').textContent = initials;
        } else {
            // Redireciona para login
            window.location.href = '/';
        }
    } catch (error) {
        console.error('Erro ao verificar autenticação:', error);
        window.location.href = '/';
    }
}

async function handleLogout() {
    try {
        await fetch('/api/auth/logout');
        window.location.href = '/';
    } catch (error) {
        console.error('Erro no logout:', error);
    }
}

function loadRecentDocuments() {
    const recentDocs = RecentDocumentsManager.get();

    const grid = document.getElementById('recentDocumentsGrid');
    const emptyState = document.getElementById('emptyState');

    if (recentDocs.length === 0) {
        grid.style.display = 'none';
        emptyState.style.display = 'flex';
    } else {
        grid.style.display = 'grid';
        emptyState.style.display = 'none';
        renderRecentDocumentsCards(recentDocs);
    }
}

function renderRecentDocumentsCards(documents) {
    const grid = document.getElementById('recentDocumentsGrid');
    grid.innerHTML = '';

    documents.forEach(doc => {
        const card = document.createElement('div');
        card.className = 'recent-card bg-white dark:bg-neutral-800 rounded-lg border border-neutral-200 dark:border-neutral-700 p-6 hover:shadow-lg transition-all';

        const fileIcon = getFileMaterialIcon(doc.type);
        const iconColor = getFileIconColor(doc.type);
        const accessedDate = doc.accessedAt ? new Date(doc.accessedAt) : new Date();

        // Se tem resumo, mostra o resumo. Senão, mostra informações do arquivo
        const contentHTML = doc.hasSummary && doc.summary ? `
            <div class="mb-4 p-4 bg-blue-50 dark:bg-blue-900/20 rounded-lg border-l-4 border-blue-500 cursor-pointer hover:bg-blue-100 dark:hover:bg-blue-900/30 transition-colors" onclick="showFullSummary('${doc.id}')">
                <div class="flex items-center gap-2 mb-2">
                    <span class="material-symbols-outlined text-blue-600 dark:text-blue-400 text-sm">auto_awesome</span>
                    <span class="text-xs font-semibold text-blue-600 dark:text-blue-400 uppercase">Resumo da IA</span>
                    <span class="ml-auto text-xs text-blue-500 dark:text-blue-400">Clique para ver completo</span>
                </div>
                <div class="text-sm text-gray-700 dark:text-gray-300 line-clamp-3">
                    ${formatMarkdownSimple(doc.summary)}
                </div>
            </div>
        ` : `
            <div class="flex items-center justify-between text-sm text-gray-500 dark:text-gray-400 mb-4">
                <div class="flex items-center gap-1">
                    <span class="material-symbols-outlined text-base">schedule</span>
                    <span>Acessado ${formatRelativeTime(accessedDate)}</span>
                </div>
                <span>${formatFileSize(doc.size || 0)}</span>
            </div>
        `;

        card.innerHTML = `
            <div class="flex items-start gap-4 mb-4">
                <span class="recent-card-icon material-symbols-outlined ${iconColor} text-5xl">${fileIcon}</span>
                <div class="flex-1 min-w-0">
                    <h3 class="font-semibold text-lg text-gray-900 dark:text-gray-100 truncate mb-1">${doc.name}</h3>
                    <p class="text-sm text-gray-500 dark:text-gray-400 truncate">${doc.driveName || 'SharePoint'}</p>
                </div>
            </div>
            ${contentHTML}
            <a href="${doc.webUrl}" target="_blank"
               class="open-doc-link block w-full text-center px-4 py-2.5 text-sm font-semibold text-white bg-primary-600 rounded-lg hover:bg-primary-700 transition-colors"
               data-doc-json='${JSON.stringify(doc).replace(/'/g, "&#39;")}'>
                <span class="flex items-center justify-center gap-2">
                    <span class="material-symbols-outlined text-lg">open_in_new</span>
                    Abrir Documento
                </span>
            </a>
        `;

        grid.appendChild(card);
    });

    // Adiciona listeners aos links de abrir
    document.querySelectorAll('.open-doc-link').forEach(link => {
        link.addEventListener('click', (e) => {
            try {
                const docJson = e.currentTarget.getAttribute('data-doc-json');
                if (docJson) {
                    const doc = JSON.parse(docJson);
                    RecentDocumentsManager.add(doc);
                }
            } catch (err) {
                console.error('Erro ao adicionar aos recentes:', err);
            }
        });
    });
}

function formatMarkdownSimple(text) {
    if (!text) return '';

    // Versão simplificada do formatMarkdown para resumos
    let html = text
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;");

    // Bold
    html = html.replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>');

    // Quebras de linha
    html = html.replace(/\n/g, '<br>');

    return html;
}

function formatRelativeTime(date) {
    const now = new Date();
    const diffMs = now - date;
    const diffMins = Math.floor(diffMs / 60000);
    const diffHours = Math.floor(diffMs / 3600000);
    const diffDays = Math.floor(diffMs / 86400000);

    if (diffMins < 1) return 'agora';
    if (diffMins < 60) return `há ${diffMins} min`;
    if (diffHours < 24) return `há ${diffHours}h`;
    if (diffDays < 7) return `há ${diffDays}d`;

    return formatDate(date.toISOString());
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

function formatFileSize(bytes) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return Math.round(bytes / Math.pow(k, i) * 100) / 100 + ' ' + sizes[i];
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

// RecentDocumentsManager (copiado do app.js para standalone)
const RecentDocumentsManager = {
    STORAGE_KEY: 'sharepoint_ai_recents',
    MAX_ITEMS: 20,

    get() {
        try {
            const stored = localStorage.getItem(this.STORAGE_KEY);
            return stored ? JSON.parse(stored) : [];
        } catch (e) {
            console.warn('Erro ao ler recentes:', e);
            return [];
        }
    },

    add(doc) {
        try {
            let recent = this.get();
            // Remove se já existe (para mover para o topo)
            recent = recent.filter(item => item.id !== doc.id);

            // Adiciona no início
            recent.unshift({
                ...doc,
                accessedAt: new Date().toISOString()
            });

            // Limita tamanho
            if (recent.length > this.MAX_ITEMS) {
                recent = recent.slice(0, this.MAX_ITEMS);
            }

            localStorage.setItem(this.STORAGE_KEY, JSON.stringify(recent));
            console.log('Documento adicionado aos recentes:', doc.name);
        } catch (e) {
            console.warn('Erro ao salvar recente:', e);
        }
    },

    clear() {
        localStorage.removeItem(this.STORAGE_KEY);
    }
};

// Função para mostrar resumo completo em modal
function showFullSummary(docId) {
    const recentDocs = RecentDocumentsManager.get();
    const doc = recentDocs.find(d => d.id === docId);

    if (!doc || !doc.summary) return;

    // Cria modal
    const modal = document.createElement('div');
    modal.className = 'fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50 p-4';
    modal.onclick = (e) => {
        if (e.target === modal) {
            modal.remove();
        }
    };

    const fileIcon = getFileMaterialIcon(doc.type);
    const iconColor = getFileIconColor(doc.type);

    modal.innerHTML = `
        <div class="bg-white dark:bg-neutral-800 rounded-lg shadow-2xl max-w-4xl w-full max-h-[90vh] overflow-hidden flex flex-col">
            <!-- Header -->
            <div class="flex items-center justify-between p-6 border-b border-neutral-200 dark:border-neutral-700">
                <div class="flex items-center gap-4">
                    <span class="material-symbols-outlined ${iconColor} text-4xl">${fileIcon}</span>
                    <div>
                        <h2 class="text-xl font-semibold text-gray-900 dark:text-gray-100">${doc.name}</h2>
                        <p class="text-sm text-gray-500 dark:text-gray-400">${doc.driveName || 'SharePoint'}</p>
                    </div>
                </div>
                <button onclick="this.closest('.fixed').remove()" class="p-2 hover:bg-neutral-100 dark:hover:bg-neutral-700 rounded-lg transition-colors">
                    <span class="material-symbols-outlined text-gray-600 dark:text-gray-400">close</span>
                </button>
            </div>
            
            <!-- Content -->
            <div class="flex-1 overflow-y-auto p-6">
                <div class="flex items-center gap-2 mb-4">
                    <span class="material-symbols-outlined text-blue-600 dark:text-blue-400">auto_awesome</span>
                    <h3 class="text-lg font-semibold text-blue-600 dark:text-blue-400">Resumo Completo da IA</h3>
                </div>
                <div class="prose dark:prose-invert max-w-none text-gray-700 dark:text-gray-300">
                    ${formatMarkdownSimple(doc.summary)}
                </div>
            </div>
            
            <!-- Footer -->
            <div class="flex items-center justify-end gap-3 p-6 border-t border-neutral-200 dark:border-neutral-700">
                <button onclick="this.closest('.fixed').remove()" class="px-4 py-2 text-sm font-medium text-gray-700 dark:text-gray-300 bg-neutral-100 dark:bg-neutral-700 rounded-lg hover:bg-neutral-200 dark:hover:bg-neutral-600 transition-colors">
                    Fechar
                </button>
                <a href="${doc.webUrl}" target="_blank" class="px-4 py-2 text-sm font-semibold text-white bg-primary-600 rounded-lg hover:bg-primary-700 transition-colors flex items-center gap-2">
                    <span class="material-symbols-outlined text-lg">open_in_new</span>
                    Abrir Documento
                </a>
            </div>
        </div>
    `;

    document.body.appendChild(modal);
}
