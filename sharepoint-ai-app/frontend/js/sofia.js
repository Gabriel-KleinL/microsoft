/**
 * Sof-IA - Chat Interface
 * Sistema de conversas e projetos com IA
 */

// ============================================
// Estado da Aplicação
// ============================================

const SofiaState = {
    currentConversation: null,
    conversations: [],
    projects: [],
    isProcessing: false,
    user: null,
    selectedModel: 'claude' // Modelo padrão
};

// ============================================
// Gerenciador de Conversas (localStorage)
// ============================================

const ConversationsManager = {
    STORAGE_KEY: 'sofia_conversations',

    getAll() {
        try {
            const stored = localStorage.getItem(this.STORAGE_KEY);
            return stored ? JSON.parse(stored) : [];
        } catch (e) {
            console.warn('Erro ao ler conversas:', e);
            return [];
        }
    },

    save(conversations) {
        try {
            localStorage.setItem(this.STORAGE_KEY, JSON.stringify(conversations));
        } catch (e) {
            console.warn('Erro ao salvar conversas:', e);
        }
    },

    create(title = 'Nova conversa') {
        const conversation = {
            id: Date.now().toString(),
            title: title,
            messages: [],
            createdAt: new Date().toISOString(),
            updatedAt: new Date().toISOString()
        };

        const conversations = this.getAll();
        conversations.unshift(conversation);
        this.save(conversations);

        return conversation;
    },

    update(conversationId, updates) {
        const conversations = this.getAll();
        const index = conversations.findIndex(c => c.id === conversationId);

        if (index !== -1) {
            conversations[index] = {
                ...conversations[index],
                ...updates,
                updatedAt: new Date().toISOString()
            };
            this.save(conversations);
            return conversations[index];
        }

        return null;
    },

    delete(conversationId) {
        const conversations = this.getAll();
        const filtered = conversations.filter(c => c.id !== conversationId);
        this.save(filtered);
    },

    addMessage(conversationId, message) {
        const conversations = this.getAll();
        const index = conversations.findIndex(c => c.id === conversationId);

        if (index !== -1) {
            conversations[index].messages.push(message);
            conversations[index].updatedAt = new Date().toISOString();

            // Auto-título da primeira mensagem
            if (conversations[index].messages.length === 2 && conversations[index].title === 'Nova conversa') {
                const firstUserMessage = conversations[index].messages.find(m => m.role === 'user');
                if (firstUserMessage) {
                    conversations[index].title = firstUserMessage.content.substring(0, 50) + (firstUserMessage.content.length > 50 ? '...' : '');
                }
            }

            this.save(conversations);
            return conversations[index];
        }

        return null;
    }
};

// ============================================
// Gerenciador de Projetos (localStorage)
// ============================================

const ProjectsManager = {
    STORAGE_KEY: 'sofia_projects',

    getAll() {
        try {
            const stored = localStorage.getItem(this.STORAGE_KEY);
            // Verifica se stored é null, undefined, ou a string "undefined"
            if (!stored || stored === 'undefined' || stored === 'null') {
                return [];
            }
            return JSON.parse(stored);
        } catch (e) {
            console.warn('Erro ao ler projetos:', e);
            // Limpa localStorage corrompido
            localStorage.removeItem(this.STORAGE_KEY);
            return [];
        }
    },

    save(projects) {
        try {
            localStorage.setItem(this.STORAGE_KEY, JSON.stringify(projects));
        } catch (e) {
            console.warn('Erro ao salvar projetos:', e);
        }
    },

    create(name, description = '', instructions = '') {
        const project = {
            id: Date.now().toString(),
            name: name,
            description: description,
            instructions: instructions,
            documents: [],
            notes: '',
            aiModel: 'claude', // Modelo padrão
            createdAt: new Date().toISOString(),
            updatedAt: new Date().toISOString()
        };

        const projects = this.getAll();
        projects.unshift(project);
        this.save(projects);

        return project;
    },

    update(projectId, updates) {
        const projects = this.getAll();
        const index = projects.findIndex(p => p.id === projectId);

        if (index !== -1) {
            projects[index] = {
                ...projects[index],
                ...updates,
                updatedAt: new Date().toISOString()
            };
            this.save(projects);
            return projects[index];
        }

        return null;
    },

    delete(projectId) {
        const projects = this.getAll();
        const filtered = projects.filter(p => p.id !== projectId);
        this.save(filtered);
    }
};

// ============================================
// Inicialização
// ============================================

document.addEventListener('DOMContentLoaded', () => {
    checkAuthStatus();
    loadConversations();
    loadProjects();
    setupEventListeners();
    setupSidebarToggles();
});

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
            SofiaState.user = {
                name: data.user_name,
                email: data.user_email
            };

            document.getElementById('userName').textContent = data.user_name;
            const initials = data.user_name.split(' ').map(n => n[0]).join('').substring(0, 2).toUpperCase();
            document.getElementById('userAvatar').textContent = initials;
        } else {
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

// ============================================
// Event Listeners
// ============================================

function setupEventListeners() {
    // Novo chat
    const newChatBtn = document.getElementById('newChatBtn');
    if (newChatBtn) {
        newChatBtn.addEventListener('click', handleNewChat);
    }

    // Tabs
    const tabConversations = document.getElementById('tabConversations');
    const tabProjects = document.getElementById('tabProjects');

    if (tabConversations && tabProjects) {
        tabConversations.addEventListener('click', () => showTab('conversations'));
        tabProjects.addEventListener('click', () => showTab('projects'));
    }

    // Chat input
    const chatInput = document.getElementById('chatInput');
    const chatSendBtn = document.getElementById('chatSendBtn');

    if (chatInput && chatSendBtn) {
        chatSendBtn.addEventListener('click', handleSendMessage);

        chatInput.addEventListener('keydown', (e) => {
            if (e.key === 'Enter' && !e.shiftKey) {
                e.preventDefault();
                handleSendMessage();
            }
        });

        // Auto-resize textarea
        chatInput.addEventListener('input', () => {
            chatInput.style.height = 'auto';
            chatInput.style.height = Math.min(chatInput.scrollHeight, 120) + 'px';
        });
    }

    // Limpar chat
    const clearChatBtn = document.getElementById('clearChatBtn');
    if (clearChatBtn) {
        clearChatBtn.addEventListener('click', handleClearChat);
    }

    // Sugestões
    document.querySelectorAll('.suggestion-card').forEach(card => {
        card.addEventListener('click', () => {
            const suggestion = card.querySelector('h3').textContent;
            document.getElementById('chatInput').value = `Ajude-me com: ${suggestion}`;
            handleSendMessage();
        });
    });

    // Novo projeto
    const newProjectBtn = document.getElementById('newProjectBtn');
    if (newProjectBtn) {
        newProjectBtn.addEventListener('click', handleNewProject);
    }
}

// ============================================
// Tabs (Conversas / Projetos)
// ============================================

function showTab(tab) {
    const tabConversations = document.getElementById('tabConversations');
    const tabProjects = document.getElementById('tabProjects');
    const conversationsPanel = document.getElementById('conversationsPanel');
    const projectsPanel = document.getElementById('projectsPanel');

    if (tab === 'conversations') {
        // Ativa tab conversas
        tabConversations.classList.add('text-primary-600', 'dark:text-primary-400', 'border-b-2', 'border-primary-600', 'dark:border-primary-400');
        tabConversations.classList.remove('text-gray-500', 'dark:text-gray-400');

        tabProjects.classList.remove('text-primary-600', 'dark:text-primary-400', 'border-b-2', 'border-primary-600', 'dark:border-primary-400');
        tabProjects.classList.add('text-gray-500', 'dark:text-gray-400');

        conversationsPanel.style.display = 'block';
        projectsPanel.style.display = 'none';
    } else {
        // Ativa tab projetos
        tabProjects.classList.add('text-primary-600', 'dark:text-primary-400', 'border-b-2', 'border-primary-600', 'dark:border-primary-400');
        tabProjects.classList.remove('text-gray-500', 'dark:text-gray-400');

        tabConversations.classList.remove('text-primary-600', 'dark:text-primary-400', 'border-b-2', 'border-primary-600', 'dark:border-primary-400');
        tabConversations.classList.add('text-gray-500', 'dark:text-gray-400');

        conversationsPanel.style.display = 'none';
        projectsPanel.style.display = 'block';
    }
}

// ============================================
// Conversas
// ============================================

function loadConversations() {
    SofiaState.conversations = ConversationsManager.getAll();
    renderConversations();
}

function renderConversations() {
    const list = document.getElementById('conversationsList');
    const emptyState = document.getElementById('conversationsEmptyState');

    if (SofiaState.conversations.length === 0) {
        list.innerHTML = '';
        emptyState.style.display = 'flex';
        return;
    }

    emptyState.style.display = 'none';
    list.innerHTML = '';

    SofiaState.conversations.forEach(conv => {
        const item = document.createElement('div');
        item.className = `conversation-item px-3 py-2 rounded-lg cursor-pointer transition-colors flex items-start justify-between gap-2 ${SofiaState.currentConversation?.id === conv.id ? 'active' : ''
            }`;

        const lastMessageTime = new Date(conv.updatedAt);
        const timeString = formatRelativeTime(lastMessageTime);

        item.innerHTML = `
            <div class="flex-1 min-w-0" data-conv-id="${conv.id}">
                <p class="text-sm font-medium truncate">${conv.title}</p>
                <p class="text-xs text-gray-500 dark:text-gray-400">${timeString}</p>
            </div>
            <button class="delete-conv p-1 hover:bg-red-100 dark:hover:bg-red-900/30 rounded transition-colors" data-conv-id="${conv.id}" title="Excluir">
                <span class="material-symbols-outlined text-red-600 dark:text-red-400 text-sm">delete</span>
            </button>
        `;

        list.appendChild(item);

        // Click para carregar conversa
        item.querySelector(`[data-conv-id="${conv.id}"]`).addEventListener('click', (e) => {
            if (!e.target.closest('.delete-conv')) {
                loadConversation(conv.id);
            }
        });

        // Deletar conversa
        item.querySelector('.delete-conv').addEventListener('click', (e) => {
            e.stopPropagation();
            handleDeleteConversation(conv.id);
        });
    });
}

function handleNewChat() {
    const conversation = ConversationsManager.create();
    SofiaState.conversations.unshift(conversation);
    SofiaState.currentConversation = conversation;

    renderConversations();
    showWelcomeState();
    document.getElementById('chatTitle').textContent = 'Sof-IA';
}

function loadConversation(conversationId) {
    const conversation = SofiaState.conversations.find(c => c.id === conversationId);

    if (!conversation) return;

    SofiaState.currentConversation = conversation;
    renderConversations();
    renderMessages();

    document.getElementById('chatTitle').textContent = conversation.title;
}

function handleDeleteConversation(conversationId) {
    if (!confirm('Deseja realmente excluir esta conversa?')) return;

    ConversationsManager.delete(conversationId);
    SofiaState.conversations = ConversationsManager.getAll();

    if (SofiaState.currentConversation?.id === conversationId) {
        SofiaState.currentConversation = null;
        showWelcomeState();
    }

    renderConversations();
}

function handleClearChat() {
    if (!SofiaState.currentConversation) return;

    if (!confirm('Deseja realmente limpar todas as mensagens desta conversa?')) return;

    SofiaState.currentConversation.messages = [];
    ConversationsManager.update(SofiaState.currentConversation.id, {
        messages: [],
        title: 'Nova conversa'
    });

    SofiaState.conversations = ConversationsManager.getAll();
    renderConversations();
    showWelcomeState();
}

// ============================================
// Mensagens
// ============================================

function showWelcomeState() {
    document.getElementById('welcomeState').classList.remove('hidden');
    document.getElementById('chatMessages').classList.add('hidden');
}

function hideWelcomeState() {
    document.getElementById('welcomeState').classList.add('hidden');
    document.getElementById('chatMessages').classList.remove('hidden');
}

function renderMessages() {
    if (!SofiaState.currentConversation) {
        showWelcomeState();
        return;
    }

    hideWelcomeState();

    const container = document.getElementById('chatMessages');
    container.innerHTML = '';

    SofiaState.currentConversation.messages.forEach(msg => {
        addMessageToUI(msg);
    });

    // Scroll para o final
    setTimeout(() => {
        const messagesArea = document.getElementById('chatMessagesArea');
        messagesArea.scrollTop = messagesArea.scrollHeight;
    }, 100);
}

function addMessageToUI(message) {
    const container = document.getElementById('chatMessages');

    const messageDiv = document.createElement('div');
    messageDiv.className = `flex gap-4 ${message.role === 'user' ? 'justify-end' : ''}`;

    if (message.role === 'assistant') {
        messageDiv.innerHTML = `
            <div class="flex-shrink-0">
                <div class="w-8 h-8 rounded-full bg-primary-100 dark:bg-primary-900 flex items-center justify-center">
                    <span class="material-symbols-outlined text-primary-600 dark:text-primary-400 text-lg">smart_toy</span>
                </div>
            </div>
            <div class="flex-1 max-w-3xl">
                <div class="bg-neutral-100 dark:bg-neutral-800 rounded-2xl px-4 py-3">
                    <div class="prose dark:prose-invert max-w-none text-sm">
                        ${formatMarkdown(message.content)}
                    </div>
                </div>
            </div>
        `;
    } else {
        messageDiv.innerHTML = `
            <div class="flex-1 max-w-3xl">
                <div class="bg-primary-600 text-white rounded-2xl px-4 py-3 ml-auto">
                    <p class="text-sm whitespace-pre-wrap">${escapeHtml(message.content)}</p>
                </div>
            </div>
        `;
    }

    container.appendChild(messageDiv);
}

async function handleSendMessage() {
    const input = document.getElementById('chatInput');
    const message = input.value.trim();

    if (!message || SofiaState.isProcessing) return;

    // Cria nova conversa se necessário
    if (!SofiaState.currentConversation) {
        const conversation = ConversationsManager.create();
        SofiaState.conversations.unshift(conversation);
        SofiaState.currentConversation = conversation;
        renderConversations();
        hideWelcomeState();
    }

    // Limpa input
    input.value = '';
    input.style.height = 'auto';

    // Adiciona mensagem do usuário
    const userMessage = {
        role: 'user',
        content: message,
        timestamp: new Date().toISOString()
    };

    SofiaState.currentConversation = ConversationsManager.addMessage(
        SofiaState.currentConversation.id,
        userMessage
    );

    addMessageToUI(userMessage);
    renderConversations();

    // Scroll para o final
    const messagesArea = document.getElementById('chatMessagesArea');
    messagesArea.scrollTop = messagesArea.scrollHeight;

    // Mostra typing indicator
    SofiaState.isProcessing = true;
    showTypingIndicator();

    try {
        // Envia para o backend
        const response = await fetch('/api/ai/chat', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            credentials: 'include',
            body: JSON.stringify({
                message: message,
                conversation_history: SofiaState.currentConversation.messages.slice(-10) // Últimas 10 mensagens
            })
        });

        if (!response.ok) {
            const errorData = await response.json().catch(() => ({}));
            throw new Error(errorData.error || 'Erro na resposta do servidor');
        }

        const data = await response.json();

        // Remove typing indicator
        removeTypingIndicator();

        // Adiciona resposta da IA
        const assistantMessage = {
            role: 'assistant',
            content: data.response,
            timestamp: new Date().toISOString()
        };

        SofiaState.currentConversation = ConversationsManager.addMessage(
            SofiaState.currentConversation.id,
            assistantMessage
        );

        addMessageToUI(assistantMessage);
        renderConversations();

        // Scroll para o final
        messagesArea.scrollTop = messagesArea.scrollHeight;

    } catch (error) {
        console.error('Erro ao enviar mensagem:', error);

        removeTypingIndicator();

        // Se erro 401, redireciona para login
        if (error.message.includes('401') || error.message.includes('Unauthorized')) {
            alert('Sua sessão expirou. Por favor, faça login novamente.');
            window.location.href = '/';
            return;
        }

        const errorMessage = {
            role: 'assistant',
            content: 'Desculpe, houve um erro ao processar sua mensagem. Por favor, tente novamente.',
            timestamp: new Date().toISOString()
        };

        SofiaState.currentConversation = ConversationsManager.addMessage(
            SofiaState.currentConversation.id,
            errorMessage
        );

        addMessageToUI(errorMessage);
    } finally {
        SofiaState.isProcessing = false;
    }
}

function showTypingIndicator() {
    const container = document.getElementById('chatMessages');

    const typingDiv = document.createElement('div');
    typingDiv.id = 'typingIndicator';
    typingDiv.className = 'flex gap-4';

    typingDiv.innerHTML = `
        <div class="flex-shrink-0">
            <div class="w-8 h-8 rounded-full bg-primary-100 dark:bg-primary-900 flex items-center justify-center">
                <span class="material-symbols-outlined text-primary-600 dark:text-primary-400 text-lg">smart_toy</span>
            </div>
        </div>
        <div class="flex-1 max-w-3xl">
            <div class="bg-neutral-100 dark:bg-neutral-800 rounded-2xl px-4 py-3">
                <div class="typing-indicator">
                    <span></span>
                    <span></span>
                    <span></span>
                </div>
            </div>
        </div>
    `;

    container.appendChild(typingDiv);

    // Scroll para o final
    const messagesArea = document.getElementById('chatMessagesArea');
    messagesArea.scrollTop = messagesArea.scrollHeight;
}

function removeTypingIndicator() {
    const indicator = document.getElementById('typingIndicator');
    if (indicator) {
        indicator.remove();
    }
}

// ============================================
// Projetos
// ============================================

function loadProjects() {
    SofiaState.projects = ProjectsManager.getAll();
    renderProjects();
}

function renderProjects() {
    const list = document.getElementById('projectsList');
    const emptyState = document.getElementById('projectsEmptyState');

    if (SofiaState.projects.length === 0) {
        list.innerHTML = '';
        emptyState.style.display = 'flex';
        return;
    }

    emptyState.style.display = 'none';
    list.innerHTML = '';

    SofiaState.projects.forEach(project => {
        const item = document.createElement('div');
        item.className = 'px-3 py-2 rounded-lg hover:bg-neutral-200 dark:hover:bg-neutral-800 cursor-pointer transition-colors';

        item.innerHTML = `
            <div class="flex items-start justify-between gap-2">
                <div class="flex-1 min-w-0">
                    <div class="flex items-center gap-2">
                        <span class="material-symbols-outlined text-primary-600 dark:text-primary-400 text-sm">folder_special</span>
                        <p class="text-sm font-medium truncate">${project.name}</p>
                    </div>
                    ${project.description ? `<p class="text-xs text-gray-500 dark:text-gray-400 mt-1 truncate">${project.description}</p>` : ''}
                    <p class="text-xs text-gray-400 dark:text-gray-500 mt-1">${project.documents.length} documentos</p>
                </div>
                <button class="delete-project p-1 hover:bg-red-100 dark:hover:bg-red-900/30 rounded transition-colors" data-project-id="${project.id}" title="Excluir">
                    <span class="material-symbols-outlined text-red-600 dark:text-red-400 text-sm">delete</span>
                </button>
            </div>
        `;

        list.appendChild(item);

        // Deletar projeto
        item.querySelector('.delete-project').addEventListener('click', (e) => {
            e.stopPropagation();
            handleDeleteProject(project.id);
        });

        // Click no projeto (abre visualização completa)
        item.addEventListener('click', (e) => {
            if (!e.target.closest('.delete-project')) {
                openProjectView(project);
            }
        });
    });
}

function handleNewProject() {
    showCreateProjectModal();
}

function showCreateProjectModal() {
    // Modal simples para criar projeto
    const modal = document.createElement('div');
    modal.className = 'fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50 p-4';
    modal.id = 'createProjectModal';

    modal.innerHTML = `
        <div class="bg-white dark:bg-neutral-800 rounded-lg shadow-2xl max-w-md w-full">
            <!-- Header -->
            <div class="flex items-center justify-between p-6 border-b border-neutral-200 dark:border-neutral-700">
                <h2 class="text-xl font-semibold text-gray-900 dark:text-gray-100">Novo Projeto</h2>
                <button onclick="document.getElementById('createProjectModal').remove()" class="p-2 hover:bg-neutral-100 dark:hover:bg-neutral-700 rounded-lg transition-colors">
                    <span class="material-symbols-outlined text-gray-600 dark:text-gray-400">close</span>
                </button>
            </div>
            
            <!-- Content -->
            <div class="p-6 space-y-4">
                <div>
                    <label class="block text-sm font-semibold text-gray-700 dark:text-gray-300 mb-2">
                        Nome do Projeto <span class="text-red-500">*</span>
                    </label>
                    <input 
                        type="text" 
                        id="newProjectName" 
                        placeholder="Ex: Análise de Vendas Q4"
                        class="w-full px-4 py-2 border border-neutral-300 dark:border-neutral-600 rounded-lg bg-white dark:bg-neutral-900 text-gray-900 dark:text-gray-100 focus:ring-2 focus:ring-primary-500 focus:border-transparent"
                        required
                    />
                </div>
                
                <div>
                    <label class="block text-sm font-semibold text-gray-700 dark:text-gray-300 mb-2">
                        Descrição
                    </label>
                    <textarea 
                        id="newProjectDescription" 
                        placeholder="Descreva brevemente o objetivo deste projeto..."
                        rows="3"
                        class="w-full px-4 py-2 border border-neutral-300 dark:border-neutral-600 rounded-lg bg-white dark:bg-neutral-900 text-gray-900 dark:text-gray-100 focus:ring-2 focus:ring-primary-500 focus:border-transparent resize-none"
                    ></textarea>
                </div>
            </div>
            
            <!-- Footer -->
            <div class="flex items-center justify-end gap-3 p-6 border-t border-neutral-200 dark:border-neutral-700">
                <button 
                    onclick="document.getElementById('createProjectModal').remove()" 
                    class="px-4 py-2 text-sm font-medium text-gray-700 dark:text-gray-300 bg-neutral-100 dark:bg-neutral-700 rounded-lg hover:bg-neutral-200 dark:hover:bg-neutral-600 transition-colors"
                >
                    Cancelar
                </button>
                <button 
                    onclick="createNewProject()" 
                    class="px-4 py-2 text-sm font-semibold text-white bg-primary-600 rounded-lg hover:bg-primary-700 transition-colors"
                >
                    Criar Projeto
                </button>
            </div>
        </div>
    `;

    document.body.appendChild(modal);
    setTimeout(() => document.getElementById('newProjectName').focus(), 100);
}

function createNewProject() {
    const name = document.getElementById('newProjectName').value.trim();
    const description = document.getElementById('newProjectDescription').value.trim();

    if (!name) {
        alert('Por favor, insira um nome para o projeto');
        document.getElementById('newProjectName').focus();
        return;
    }

    const project = ProjectsManager.create(name, description);
    SofiaState.projects.unshift(project);
    renderProjects();

    document.getElementById('createProjectModal').remove();

    // Abre o projeto recém-criado
    openProjectView(project);
}

function openProjectView(project) {
    // Esconde a interface principal
    document.querySelector('.flex.h-screen').style.display = 'none';

    // Cria visualização completa do projeto
    const projectView = document.createElement('div');
    projectView.id = 'projectView';
    projectView.className = 'flex h-screen bg-neutral-50 dark:bg-neutral-900';

    projectView.innerHTML = `
        <!-- Sidebar Esquerda - Conversas do Projeto -->
        <div class="w-80 bg-white dark:bg-neutral-800 border-r border-neutral-200 dark:border-neutral-700 flex flex-col">
            <div class="p-4 border-b border-neutral-200 dark:border-neutral-700">
                <button onclick="closeProjectView()" class="flex items-center gap-2 text-sm text-gray-600 dark:text-gray-400 hover:text-gray-900 dark:hover:text-gray-100 mb-4">
                    <span class="material-symbols-outlined text-lg">arrow_back</span>
                    Todos os projetos
                </button>
                <h2 class="text-lg font-semibold text-gray-900 dark:text-gray-100 truncate">${project.name}</h2>
                ${project.description ? `<p class="text-sm text-gray-500 dark:text-gray-400 mt-1">${project.description}</p>` : ''}
            </div>
            
            <div class="flex-1 overflow-y-auto p-4">
                <div class="flex items-center justify-between mb-3">
                    <h3 class="text-sm font-semibold text-gray-700 dark:text-gray-300">Conversas</h3>
                    <button onclick="newProjectConversation('${project.id}')" class="p-1 hover:bg-neutral-100 dark:hover:bg-neutral-700 rounded transition-colors">
                        <span class="material-symbols-outlined text-gray-600 dark:text-gray-400 text-lg">add</span>
                    </button>
                </div>
                <div id="projectConversationsList" class="space-y-2">
                    <!-- Conversas do projeto -->
                </div>
            </div>
        </div>
        
        <!-- Área Principal - Chat -->
        <div class="flex-1 flex flex-col">
            <!-- Header do Chat -->
            <div class="bg-white dark:bg-neutral-800 border-b border-neutral-200 dark:border-neutral-700 p-4 flex items-center justify-between">
                <h3 class="text-lg font-semibold text-gray-900 dark:text-gray-100">Chat do Projeto</h3>
                <button onclick="clearProjectChat('${project.id}')" class="text-sm text-gray-600 dark:text-gray-400 hover:text-gray-900 dark:hover:text-gray-100 flex items-center gap-1">
                    <span class="material-symbols-outlined text-lg">delete_sweep</span>
                    Limpar
                </button>
            </div>
            
            <!-- Mensagens -->
            <div id="projectChatMessagesArea" class="flex-1 overflow-y-auto p-6">
                <div id="projectChatMessages" class="max-w-4xl mx-auto space-y-6">
                    <!-- Welcome state -->
                    <div id="projectWelcomeState" class="text-center py-12">
                        <span class="material-symbols-outlined text-primary-600 dark:text-primary-400 text-6xl mb-4">folder_special</span>
                        <h3 class="text-xl font-semibold text-gray-900 dark:text-gray-100 mb-2">Bem-vindo ao ${project.name}</h3>
                        <p class="text-gray-600 dark:text-gray-400">Faça perguntas sobre este projeto e use os arquivos de referência</p>
                    </div>
                </div>
            </div>
            
            <!-- Input -->
            <div class="bg-white dark:bg-neutral-800 border-t border-neutral-200 dark:border-neutral-700 p-4">
                <div class="max-w-4xl mx-auto">
                    <div class="flex gap-3">
                        <textarea 
                            id="projectChatInput" 
                            placeholder="Faça uma pergunta sobre este projeto..."
                            rows="1"
                            class="flex-1 px-4 py-3 border border-neutral-300 dark:border-neutral-600 rounded-lg bg-white dark:bg-neutral-900 text-gray-900 dark:text-gray-100 focus:ring-2 focus:ring-primary-500 focus:border-transparent resize-none"
                        ></textarea>
                        <button 
                            id="projectChatSendBtn"
                            class="px-6 py-3 bg-primary-600 text-white rounded-lg hover:bg-primary-700 transition-colors flex items-center gap-2"
                        >
                            <span class="material-symbols-outlined">send</span>
                        </button>
                    </div>
                </div>
            </div>
        </div>
        
        <!-- Sidebar Direita - Instruções e Arquivos -->
        <div class="w-80 bg-white dark:bg-neutral-800 border-l border-neutral-200 dark:border-neutral-700 flex flex-col overflow-y-auto">
            <!-- Instruções -->
            <div class="p-4 border-b border-neutral-200 dark:border-neutral-700">
                <div class="flex items-center justify-between mb-3">
                    <h3 class="text-sm font-semibold text-gray-700 dark:text-gray-300">Instruções</h3>
                    <button onclick="editProjectInstructions('${project.id}')" class="text-xs text-primary-600 dark:text-primary-400 hover:underline">
                        ${project.instructions ? 'Editar' : 'Adicionar'}
                    </button>
                </div>
                <div id="projectInstructionsDisplay" class="text-sm text-gray-600 dark:text-gray-400">
                    ${project.instructions || '<p class="text-gray-400 dark:text-gray-500 italic">Nenhuma instrução definida</p>'}
                </div>
            </div>
            
            <!-- Modelo de IA -->
            <div class="p-4 border-b border-neutral-200 dark:border-neutral-700">
                <div class="mb-3">
                    <h3 class="text-sm font-semibold text-gray-700 dark:text-gray-300 mb-2">Modelo de IA</h3>
                    <p class="text-xs text-gray-500 dark:text-gray-400 mb-3">Escolha qual IA usar neste projeto</p>
                </div>
                <div class="flex gap-2">
                    <button 
                        onclick="setProjectAIModel('${project.id}', 'claude')" 
                        class="flex-1 px-3 py-2 rounded-lg text-xs font-medium transition-all ${(project.aiModel || 'claude') === 'claude' ? 'bg-primary-600 text-white' : 'bg-neutral-100 dark:bg-neutral-700 text-gray-700 dark:text-gray-300 hover:bg-neutral-200 dark:hover:bg-neutral-600'}"
                        id="modelBtn_${project.id}_claude"
                    >
                        Claude
                    </button>
                    <button 
                        onclick="setProjectAIModel('${project.id}', 'gpt')" 
                        class="flex-1 px-3 py-2 rounded-lg text-xs font-medium transition-all ${(project.aiModel || 'claude') === 'gpt' ? 'bg-primary-600 text-white' : 'bg-neutral-100 dark:bg-neutral-700 text-gray-700 dark:text-gray-300 hover:bg-neutral-200 dark:hover:bg-neutral-600'}"
                        id="modelBtn_${project.id}_gpt"
                    >
                        GPT
                    </button>
                </div>
            </div>
            
            <!-- Arquivos -->
            <div class="p-4">
                <div class="flex items-center justify-between mb-3">
                    <h3 class="text-sm font-semibold text-gray-700 dark:text-gray-300">Arquivos</h3>
                    <button onclick="addProjectFiles('${project.id}')" class="text-xs text-primary-600 dark:text-primary-400 hover:underline flex items-center gap-1">
                        <span class="material-symbols-outlined text-sm">add</span>
                        Adicionar
                    </button>
                </div>
                <div id="projectFilesDisplay" class="space-y-2">
                    ${renderProjectFilesList(project.documents || [])}
                </div>
            </div>
        </div>
    `;

    document.body.appendChild(projectView);

    // Setup project chat
    setupProjectChat(project.id);
}

function renderProjectFilesList(documents) {
    if (!documents || documents.length === 0) {
        return `
            <div class="text-center py-6 border-2 border-dashed border-neutral-300 dark:border-neutral-600 rounded-lg">
                <span class="material-symbols-outlined text-gray-400 text-3xl mb-2">description</span>
                <p class="text-xs text-gray-500 dark:text-gray-400">
                    Nenhum arquivo adicionado
                </p>
            </div>
        `;
    }

    return documents.map(doc => {
        const isFolder = doc.type === 'folder' || doc.isFolder;
        const icon = isFolder ? 'folder' : 'description';
        const iconColor = isFolder ? 'text-yellow-600 dark:text-yellow-400' : 'text-primary-600 dark:text-primary-400';

        return `
        <div class="flex items-center gap-2 p-2 bg-neutral-50 dark:bg-neutral-900 rounded-lg border border-neutral-200 dark:border-neutral-700">
            <span class="material-symbols-outlined ${iconColor} text-sm">${icon}</span>
            <div class="flex-1 min-w-0">
                <p class="text-xs font-medium text-gray-900 dark:text-gray-100 truncate" title="${doc.name}">${doc.name}</p>
            </div>
            <button onclick="removeProjectFile('${doc.id}')" class="p-1 hover:bg-red-100 dark:hover:bg-red-900/30 rounded transition-colors">
                <span class="material-symbols-outlined text-red-600 dark:text-red-400 text-sm">close</span>
            </button>
        </div>
    `}).join('');
}

function setupProjectChat(projectId) {
    const input = document.getElementById('projectChatInput');
    const sendBtn = document.getElementById('projectChatSendBtn');

    if (input && sendBtn) {
        sendBtn.addEventListener('click', () => sendProjectMessage(projectId));

        input.addEventListener('keydown', (e) => {
            if (e.key === 'Enter' && !e.shiftKey) {
                e.preventDefault();
                sendProjectMessage(projectId);
            }
        });

        input.addEventListener('input', () => {
            input.style.height = 'auto';
            input.style.height = Math.min(input.scrollHeight, 120) + 'px';
        });
    }
}

function sendProjectMessage(projectId) {
    const input = document.getElementById('projectChatInput');
    const message = input.value.trim();

    if (!message) return;

    // TODO: Implementar envio de mensagem do projeto
    console.log('Enviar mensagem do projeto:', projectId, message);

    input.value = '';
    input.style.height = 'auto';
}

function closeProjectView() {
    const projectView = document.getElementById('projectView');
    if (projectView) {
        projectView.remove();
    }
    document.querySelector('.flex.h-screen').style.display = 'flex';
}

function newProjectConversation(projectId) {
    // TODO: Criar nova conversa no projeto
    console.log('Nova conversa no projeto:', projectId);
}

function clearProjectChat(projectId) {
    if (confirm('Deseja limpar todas as mensagens deste chat?')) {
        // TODO: Limpar chat do projeto
        console.log('Limpar chat do projeto:', projectId);
    }
}

function editProjectInstructions(projectId) {
    const project = SofiaState.projects.find(p => p.id === projectId);
    if (!project) return;

    showEditInstructionsModal(project);
}

function showEditInstructionsModal(project) {
    const modal = document.createElement('div');
    modal.className = 'fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50 p-4';
    modal.id = 'editInstructionsModal';

    modal.innerHTML = `
        <div class="bg-white dark:bg-neutral-800 rounded-lg shadow-2xl max-w-2xl w-full flex flex-col max-h-[90vh]">
            <!-- Header -->
            <div class="flex items-center justify-between p-6 border-b border-neutral-200 dark:border-neutral-700 shrink-0">
                <h2 class="text-xl font-semibold text-gray-900 dark:text-gray-100">Instruções do Projeto</h2>
                <button onclick="document.getElementById('editInstructionsModal').remove()" class="p-2 hover:bg-neutral-100 dark:hover:bg-neutral-700 rounded-lg transition-colors">
                    <span class="material-symbols-outlined text-gray-600 dark:text-gray-400">close</span>
                </button>
            </div>
            
            <!-- Content -->
            <div class="p-6 flex-1 overflow-y-auto">
                <p class="text-sm text-gray-600 dark:text-gray-400 mb-4">
                    Defina como a IA deve se comportar neste projeto. Você pode especificar tom de voz, formato de resposta, ou regras específicas.
                </p>
                <textarea 
                    id="projectInstructionsInput" 
                    placeholder="Ex: Aja como um especialista em vendas. Responda sempre com bullet points. Use tom formal."
                    class="w-full h-64 px-4 py-3 border border-neutral-300 dark:border-neutral-600 rounded-lg bg-white dark:bg-neutral-900 text-gray-900 dark:text-gray-100 focus:ring-2 focus:ring-primary-500 focus:border-transparent resize-none font-mono text-sm"
                >${project.instructions || ''}</textarea>
            </div>
            
            <!-- Footer -->
            <div class="flex items-center justify-end gap-3 p-6 border-t border-neutral-200 dark:border-neutral-700 shrink-0">
                <button 
                    onclick="document.getElementById('editInstructionsModal').remove()" 
                    class="px-4 py-2 text-sm font-medium text-gray-700 dark:text-gray-300 bg-neutral-100 dark:bg-neutral-700 rounded-lg hover:bg-neutral-200 dark:hover:bg-neutral-600 transition-colors"
                >
                    Cancelar
                </button>
                <button 
                    onclick="saveProjectInstructions('${project.id}')" 
                    class="px-4 py-2 text-sm font-semibold text-white bg-primary-600 rounded-lg hover:bg-primary-700 transition-colors"
                >
                    Salvar Instruções
                </button>
            </div>
        </div>
    `;

    document.body.appendChild(modal);
    setTimeout(() => document.getElementById('projectInstructionsInput').focus(), 100);
}

function saveProjectInstructions(projectId) {
    const instructions = document.getElementById('projectInstructionsInput').value.trim();
    const project = SofiaState.projects.find(p => p.id === projectId);

    if (project) {
        project.instructions = instructions;
        ProjectsManager.save(SofiaState.projects);

        // Atualiza UI
        const display = document.getElementById('projectInstructionsDisplay');
        if (display) {
            display.innerHTML = instructions || '<p class="text-gray-400 dark:text-gray-500 italic">Nenhuma instrução definida</p>';
        }

        // Atualiza botão
        const btn = document.querySelector(`button[onclick="editProjectInstructions('${projectId}')"]`);
        if (btn) {
            btn.textContent = instructions ? 'Editar' : 'Adicionar';
        }
    }

    document.getElementById('editInstructionsModal').remove();
}
function addProjectFiles(projectId) {
    const modal = document.createElement('div');
    modal.className = 'fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50 p-4';
    modal.id = 'addFileModal';

    modal.innerHTML = `
        <div class="bg-white dark:bg-neutral-800 rounded-lg shadow-2xl max-w-md w-full">
            <div class="flex items-center justify-between p-6 border-b border-neutral-200 dark:border-neutral-700">
                <h2 class="text-xl font-semibold text-gray-900 dark:text-gray-100">Adicionar Arquivos</h2>
                <button onclick="document.getElementById('addFileModal').remove()" class="p-2 hover:bg-neutral-100 dark:hover:bg-neutral-700 rounded-lg transition-colors">
                    <span class="material-symbols-outlined text-gray-600 dark:text-gray-400">close</span>
                </button>
            </div>
            
            <div class="p-6 space-y-4">
                <button onclick="document.getElementById('localFileInput').click()" class="w-full flex items-center gap-4 p-4 rounded-lg border-2 border-neutral-200 dark:border-neutral-700 hover:border-primary-500 dark:hover:border-primary-500 hover:bg-primary-50 dark:hover:bg-primary-900/20 transition-all group text-left">
                    <div class="w-12 h-12 rounded-full bg-primary-100 dark:bg-primary-900/50 flex items-center justify-center group-hover:bg-primary-200 dark:group-hover:bg-primary-800 transition-colors">
                        <span class="material-symbols-outlined text-primary-600 dark:text-primary-400 text-2xl">upload_file</span>
                    </div>
                    <div>
                        <h3 class="font-semibold text-gray-900 dark:text-gray-100">Meu Computador</h3>
                        <p class="text-xs text-gray-500 dark:text-gray-400">Faça upload de arquivos locais</p>
                    </div>
                </button>
                
                <button onclick="showSharePointFilePicker('${projectId}')" class="w-full flex items-center gap-4 p-4 rounded-lg border-2 border-neutral-200 dark:border-neutral-700 hover:border-primary-500 dark:hover:border-primary-500 hover:bg-primary-50 dark:hover:bg-primary-900/20 transition-all group text-left">
                    <div class="w-12 h-12 rounded-full bg-blue-100 dark:bg-blue-900/50 flex items-center justify-center group-hover:bg-blue-200 dark:group-hover:bg-blue-800 transition-colors">
                        <span class="material-symbols-outlined text-blue-600 dark:text-blue-400 text-2xl">cloud_download</span>
                    </div>
                    <div>
                        <h3 class="font-semibold text-gray-900 dark:text-gray-100">SharePoint</h3>
                        <p class="text-xs text-gray-500 dark:text-gray-400">Selecione arquivos da nuvem</p>
                    </div>
                </button>
            </div>
        </div>
        <input type="file" id="localFileInput" class="hidden" multiple onchange="handleLocalFileUpload(this, '${projectId}')">
    `;

    document.body.appendChild(modal);
}

function handleLocalFileUpload(input, projectId) {
    const files = Array.from(input.files);
    if (files.length === 0) return;

    const project = SofiaState.projects.find(p => p.id === projectId);
    if (!project) return;

    // Simula upload (por enquanto apenas adiciona referência local)
    files.forEach(file => {
        const doc = {
            id: 'local_' + Date.now() + '_' + Math.random().toString(36).substr(2, 9),
            name: file.name,
            size: formatFileSize(file.size),
            type: file.type || 'application/octet-stream',
            source: 'local',
            lastModified: new Date(file.lastModified).toISOString()
        };

        if (!project.documents) project.documents = [];
        project.documents.push(doc);
    });

    ProjectsManager.save(SofiaState.projects);

    // Atualiza UI
    const filesList = document.getElementById('projectFilesDisplay');
    if (filesList) {
        filesList.innerHTML = renderProjectFilesList(project.documents);
    }

    document.getElementById('addFileModal').remove();
    showToast(`${files.length} arquivo(s) adicionado(s) com sucesso!`, 'success');
}

async function showSharePointFilePicker(projectId) {
    // Fecha modal anterior
    const prevModal = document.getElementById('addFileModal');
    if (prevModal) prevModal.remove();

    // Cria modal
    const modal = document.createElement('div');
    modal.className = 'fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50 p-4';
    modal.id = 'sharePointPickerModal';

    modal.innerHTML = `
        <div class="bg-white dark:bg-neutral-800 rounded-lg shadow-2xl max-w-2xl w-full flex flex-col max-h-[80vh]">
            <!-- Header -->
            <div class="flex items-center justify-between p-6 border-b border-neutral-200 dark:border-neutral-700 shrink-0">
                <h2 class="text-xl font-semibold text-gray-900 dark:text-gray-100">Selecionar do SharePoint</h2>
                <button onclick="document.getElementById('sharePointPickerModal').remove()" class="p-2 hover:bg-neutral-100 dark:hover:bg-neutral-700 rounded-lg transition-colors">
                    <span class="material-symbols-outlined text-gray-600 dark:text-gray-400">close</span>
                </button>
            </div>
            
            <!-- Loading State -->
            <div id="spPickerLoading" class="p-12 flex flex-col items-center justify-center">
                <div class="animate-spin rounded-full h-10 w-10 border-b-2 border-primary-600 mb-4"></div>
                <p class="text-gray-500 dark:text-gray-400">Carregando documentos...</p>
            </div>
            
            <!-- Content (Lista) -->
            <div id="spPickerContent" class="hidden flex-1 overflow-y-auto p-2">
                <div id="spPickerList" class="space-y-1"></div>
            </div>
            
            <!-- Footer -->
            <div class="flex items-center justify-between p-6 border-t border-neutral-200 dark:border-neutral-700 shrink-0">
                <span id="spPickerCount" class="text-sm text-gray-500 dark:text-gray-400">0 selecionado(s)</span>
                <div class="flex gap-3">
                    <button 
                        onclick="document.getElementById('sharePointPickerModal').remove()" 
                        class="px-4 py-2 text-sm font-medium text-gray-700 dark:text-gray-300 bg-neutral-100 dark:bg-neutral-700 rounded-lg hover:bg-neutral-200 dark:hover:bg-neutral-600 transition-colors"
                    >
                        Cancelar
                    </button>
                    <button 
                        onclick="addSelectedSharePointFiles('${projectId}')" 
                        class="px-4 py-2 text-sm font-semibold text-white bg-primary-600 rounded-lg hover:bg-primary-700 transition-colors"
                    >
                        Adicionar Selecionados
                    </button>
                </div>
            </div>
        </div>
    `;

    document.body.appendChild(modal);

    // Carrega documentos
    try {
        const response = await fetch('/api/documents/list');
        const data = await response.json();

        const listContainer = document.getElementById('spPickerList');
        const loading = document.getElementById('spPickerLoading');
        const content = document.getElementById('spPickerContent');

        loading.classList.add('hidden');
        content.classList.remove('hidden');

        if (data.items && data.items.length > 0) {
            // Mostra tudo (arquivos e pastas)
            listContainer.innerHTML = data.items.map(item => {
                const isFolder = item.folder || item.isFolder;
                const icon = isFolder ? 'folder' : 'description';
                const iconColor = isFolder ? 'text-yellow-600 dark:text-yellow-400' : 'text-blue-600 dark:text-blue-400';
                const bgColor = isFolder ? 'bg-yellow-50 dark:bg-yellow-900/20' : 'bg-blue-50 dark:bg-blue-900/20';
                const sizeText = isFolder ? (item.folder?.childCount + ' itens' || 'Pasta') : formatFileSize(item.size);

                return `
                <label class="flex items-center gap-3 p-3 hover:bg-neutral-50 dark:hover:bg-neutral-700/50 rounded-lg cursor-pointer border border-transparent hover:border-neutral-200 dark:hover:border-neutral-700 transition-all">
                    <input type="checkbox" value='${JSON.stringify(item)}' class="sp-file-checkbox w-5 h-5 rounded border-gray-300 text-primary-600 focus:ring-primary-500" onchange="updatePickerCount()">
                    <div class="w-10 h-10 rounded-lg ${bgColor} flex items-center justify-center shrink-0">
                        <span class="material-symbols-outlined ${iconColor}">${icon}</span>
                    </div>
                    <div class="flex-1 min-w-0">
                        <p class="font-medium text-gray-900 dark:text-gray-100 truncate">${item.name}</p>
                        <p class="text-xs text-gray-500 dark:text-gray-400">
                            ${sizeText} • ${new Date(item.lastModifiedDateTime).toLocaleDateString()}
                        </p>
                    </div>
                </label>
            `}).join('');
        } else {
            listContainer.innerHTML = `
                <div class="text-center py-12">
                    <p class="text-gray-500 dark:text-gray-400">Nenhum arquivo encontrado.</p>
                </div>
            `;
        }
    } catch (error) {
        console.error('Erro ao carregar documentos:', error);
        document.getElementById('spPickerLoading').innerHTML = `
            <div class="text-red-500 text-center">
                <span class="material-symbols-outlined text-4xl mb-2">error</span>
                <p>Erro ao carregar documentos.</p>
            </div>
        `;
    }
}

function updatePickerCount() {
    const checked = document.querySelectorAll('.sp-file-checkbox:checked').length;
    document.getElementById('spPickerCount').textContent = `${checked} selecionado(s)`;
}

function addSelectedSharePointFiles(projectId) {
    const checkboxes = document.querySelectorAll('.sp-file-checkbox:checked');
    const project = SofiaState.projects.find(p => p.id === projectId);

    if (!project || checkboxes.length === 0) return;

    if (!project.documents) project.documents = [];

    let addedCount = 0;
    checkboxes.forEach(cb => {
        const file = JSON.parse(cb.value);

        // Evita duplicatas
        if (!project.documents.some(d => d.id === file.id)) {
            const isFolder = file.folder || file.isFolder;

            project.documents.push({
                id: file.id,
                name: file.name,
                size: isFolder ? (file.folder?.childCount + ' itens' || 'Pasta') : formatFileSize(file.size),
                type: isFolder ? 'folder' : 'sharepoint', // Marca como pasta ou arquivo
                source: 'sharepoint',
                webUrl: file.webUrl,
                downloadUrl: file['@microsoft.graph.downloadUrl'],
                lastModified: file.lastModifiedDateTime,
                isFolder: !!isFolder
            });
            addedCount++;
        }
    });

    if (addedCount > 0) {
        ProjectsManager.save(SofiaState.projects);

        // Atualiza UI
        const filesList = document.getElementById('projectFilesDisplay');
        if (filesList) {
            filesList.innerHTML = renderProjectFilesList(project.documents);
        }

        showToast(`${addedCount} arquivo(s) adicionado(s)!`, 'success');
    } else {
        showToast('Arquivos selecionados já estão no projeto.', 'info');
    }

    document.getElementById('sharePointPickerModal').remove();
}

function removeProjectFile(fileId) {
    // TODO: Implementar remoção de arquivo
    console.log('Remover arquivo:', fileId);
}

function setProjectAIModel(projectId, model) {
    const project = SofiaState.projects.find(p => p.id === projectId);

    if (project) {
        project.aiModel = model;
        ProjectsManager.save(SofiaState.projects);

        // Atualiza UI dos botões
        const claudeBtn = document.getElementById(`modelBtn_${projectId}_claude`);
        const gptBtn = document.getElementById(`modelBtn_${projectId}_gpt`);

        if (claudeBtn && gptBtn) {
            // Remove classes ativas
            claudeBtn.className = 'flex-1 px-3 py-2 rounded-lg text-xs font-medium transition-all bg-neutral-100 dark:bg-neutral-700 text-gray-700 dark:text-gray-300 hover:bg-neutral-200 dark:hover:bg-neutral-600';
            gptBtn.className = 'flex-1 px-3 py-2 rounded-lg text-xs font-medium transition-all bg-neutral-100 dark:bg-neutral-700 text-gray-700 dark:text-gray-300 hover:bg-neutral-200 dark:hover:bg-neutral-600';

            // Adiciona classe ativa ao botão selecionado
            const activeBtn = model === 'claude' ? claudeBtn : gptBtn;
            activeBtn.className = 'flex-1 px-3 py-2 rounded-lg text-xs font-medium transition-all bg-primary-600 text-white';
        }

        showToast(`Modelo alterado para ${model === 'claude' ? 'Claude' : 'GPT'}`, 'success');
    }
}

function formatFileSize(bytes) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return Math.round(bytes / Math.pow(k, i) * 100) / 100 + ' ' + sizes[i];
}

function handleDeleteProject(projectId) {
    if (!confirm('Deseja realmente excluir este projeto?')) return;

    ProjectsManager.delete(projectId);
    SofiaState.projects = ProjectsManager.getAll();
    renderProjects();
}

// ============================================
// Utilitários
// ============================================

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

    return date.toLocaleDateString('pt-BR', {
        day: '2-digit',
        month: '2-digit'
    });
}

function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

function formatMarkdown(text) {
    if (!text) return '';

    let html = escapeHtml(text);

    // Bold
    html = html.replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>');

    // Italic
    html = html.replace(/\*(.*?)\*/g, '<em>$1</em>');

    // Code blocks
    html = html.replace(/```([\s\S]*?)```/g, '<pre class="bg-neutral-200 dark:bg-neutral-900 p-3 rounded-lg overflow-x-auto"><code>$1</code></pre>');

    // Inline code
    html = html.replace(/`([^`]+)`/g, '<code class="bg-neutral-200 dark:bg-neutral-900 px-1.5 py-0.5 rounded text-xs">$1</code>');

    // Links
    html = html.replace(/\[([^\]]+)\]\(([^\)]+)\)/g, '<a href="$2" target="_blank" class="text-primary-600 dark:text-primary-400 underline">$1</a>');

    // Quebras de linha
    html = html.replace(/\n/g, '<br>');

    return html;
}

// ============================================
// Sidebar Toggles (Colapsar/Expandir)
// ============================================

function setupSidebarToggles() {
    const mainSidebar = document.getElementById('mainSidebar');
    const conversationsSidebar = document.getElementById('conversationsSidebar');
    const toggleMainBtn = document.getElementById('toggleMainSidebar');
    const toggleConversationsBtn = document.getElementById('toggleConversationsSidebar');

    let mainSidebarCollapsed = false;
    let conversationsSidebarCollapsed = false;

    // Toggle main sidebar
    if (toggleMainBtn) {
        toggleMainBtn.addEventListener('click', () => {
            mainSidebarCollapsed = !mainSidebarCollapsed;

            if (mainSidebarCollapsed) {
                mainSidebar.classList.remove('w-64');
                mainSidebar.classList.add('w-16');
                // Oculta textos
                document.querySelectorAll('.main-sidebar-text').forEach(el => {
                    el.style.display = 'none';
                });
                // Inverte ícone
                toggleMainBtn.querySelector('.material-symbols-outlined').textContent = 'chevron_right';
            } else {
                mainSidebar.classList.remove('w-16');
                mainSidebar.classList.add('w-64');
                // Mostra textos
                document.querySelectorAll('.main-sidebar-text').forEach(el => {
                    el.style.display = '';
                });
                // Inverte ícone
                toggleMainBtn.querySelector('.material-symbols-outlined').textContent = 'chevron_left';
            }
        });
    }

    // Toggle conversations sidebar
    if (toggleConversationsBtn) {
        toggleConversationsBtn.addEventListener('click', () => {
            conversationsSidebarCollapsed = !conversationsSidebarCollapsed;

            if (conversationsSidebarCollapsed) {
                conversationsSidebar.classList.remove('w-72');
                conversationsSidebar.classList.add('w-0', 'overflow-hidden');
                // Inverte ícone
                toggleConversationsBtn.querySelector('.material-symbols-outlined').textContent = 'chevron_right';
            } else {
                conversationsSidebar.classList.remove('w-0', 'overflow-hidden');
                conversationsSidebar.classList.add('w-72');
                // Inverte ícone
                toggleConversationsBtn.querySelector('.material-symbols-outlined').textContent = 'chevron_left';
            }
        });
    }
}

// ============================================
// Seletor de Modelo (Claude/GPT)
// ============================================

function selectModel(model) {
    SofiaState.selectedModel = model;

    // Atualiza visual dos botões
    const claudeBtn = document.getElementById('modelClaude');
    const gptBtn = document.getElementById('modelGPT');

    if (model === 'claude') {
        claudeBtn.className = 'model-selector px-3 py-1 text-xs font-medium rounded-lg bg-primary-600 text-white';
        gptBtn.className = 'model-selector px-3 py-1 text-xs font-medium rounded-lg bg-neutral-200 dark:bg-neutral-700 text-gray-700 dark:text-gray-300';
    } else {
        claudeBtn.className = 'model-selector px-3 py-1 text-xs font-medium rounded-lg bg-neutral-200 dark:bg-neutral-700 text-gray-700 dark:text-gray-300';
        gptBtn.className = 'model-selector px-3 py-1 text-xs font-medium rounded-lg bg-primary-600 text-white';
    }

    console.log('Modelo selecionado:', model);
}
