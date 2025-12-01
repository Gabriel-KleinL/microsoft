/**
 * Gerenciador de Configurações Compartilhado
 */

const SettingsManager = {
    STORAGE_KEY: 'sharepoint_ai_settings',
    defaults: {
        summaryLevel: 'medium' // low, medium, high
    },

    get() {
        const settings = localStorage.getItem(this.STORAGE_KEY);
        return settings ? { ...this.defaults, ...JSON.parse(settings) } : this.defaults;
    },

    set(key, value) {
        const settings = this.get();
        settings[key] = value;
        localStorage.setItem(this.STORAGE_KEY, JSON.stringify(settings));
        this.updateUI();
    },

    setSummaryLevel(level) {
        this.set('summaryLevel', level);
    },

    updateUI() {
        const settings = this.get();

        // Atualiza botões de nível de resumo
        ['low', 'medium', 'high'].forEach(level => {
            const btn = document.getElementById(`summaryLevel_${level}`);
            if (btn) {
                if (settings.summaryLevel === level) {
                    btn.classList.add('border-primary-600', 'bg-primary-50', 'text-primary-700', 'dark:border-primary-500', 'dark:bg-primary-900/20', 'dark:text-primary-300');
                    btn.classList.remove('border-neutral-200', 'dark:border-neutral-700', 'text-gray-600', 'dark:text-gray-400');
                } else {
                    btn.classList.remove('border-primary-600', 'bg-primary-50', 'text-primary-700', 'dark:border-primary-500', 'dark:bg-primary-900/20', 'dark:text-primary-300');
                    btn.classList.add('border-neutral-200', 'dark:border-neutral-700', 'text-gray-600', 'dark:text-gray-400');
                }
            }
        });
    },

    init() {
        this.updateUI();
    }
};

function openSettingsModal() {
    const modal = document.getElementById('settingsModal');
    if (modal) {
        modal.classList.remove('hidden');
        SettingsManager.updateUI();
    }
}

function closeSettingsModal() {
    const modal = document.getElementById('settingsModal');
    if (modal) {
        modal.classList.add('hidden');
    }
}

// Inicializa se o DOM já estiver carregado, ou aguarda
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', () => SettingsManager.init());
} else {
    SettingsManager.init();
}
