"""
Configurações da aplicação SharePoint + IA
"""
import os
from dotenv import load_dotenv

# Carrega variáveis de ambiente do arquivo .env
load_dotenv()

class Config:
    """Classe de configuração para gerenciar variáveis de ambiente"""

    # Microsoft Azure AD / Graph API
    MICROSOFT_CLIENT_ID = os.getenv('MICROSOFT_CLIENT_ID')
    MICROSOFT_CLIENT_SECRET = os.getenv('MICROSOFT_CLIENT_SECRET')
    MICROSOFT_TENANT_ID = os.getenv('MICROSOFT_TENANT_ID')
    MICROSOFT_REDIRECT_URI = os.getenv('MICROSOFT_REDIRECT_URI', 'http://localhost:5000/auth/callback')

    # SharePoint
    SHAREPOINT_SITE_URL = os.getenv('SHAREPOINT_SITE_URL')

    # Claude API (Anthropic)
    CLAUDE_API_KEY = os.getenv('CLAUDE_API_KEY')
    CLAUDE_MODEL = os.getenv('CLAUDE_MODEL', 'claude-sonnet-4-5-20250929')

    # OpenAI API
    OPENAI_API_KEY = os.getenv('OPENAI_API_KEY')
    OPENAI_MODEL = os.getenv('OPENAI_MODEL', 'gpt-4')

    # Provedor de IA (claude ou openai)
    AI_PROVIDER = os.getenv('AI_PROVIDER', 'claude').lower()

    # Flask
    SECRET_KEY = os.getenv('SECRET_KEY', 'dev-secret-key-change-in-production')
    FLASK_ENV = os.getenv('FLASK_ENV', 'development')

    # Scopes necessários para Microsoft Graph
    SCOPES = [
        'User.Read',
        'Files.Read.All',
        'Sites.Read.All'
    ]

    @staticmethod
    def validate():
        """Valida se todas as configurações necessárias estão presentes"""
        required_base = [
            'MICROSOFT_CLIENT_ID',
            'MICROSOFT_CLIENT_SECRET',
            'MICROSOFT_TENANT_ID',
            'SHAREPOINT_SITE_URL'
        ]

        # Valida provedor de IA
        ai_provider = Config.AI_PROVIDER
        if ai_provider not in ['claude', 'openai']:
            raise ValueError(f"AI_PROVIDER inválido: {ai_provider}. Use 'claude' ou 'openai'")

        # Adiciona chave de API necessária baseada no provedor
        if ai_provider == 'claude':
            required_base.append('CLAUDE_API_KEY')
        elif ai_provider == 'openai':
            required_base.append('OPENAI_API_KEY')

        missing = []
        for key in required_base:
            if not getattr(Config, key):
                missing.append(key)

        if missing:
            raise ValueError(f"Configurações ausentes: {', '.join(missing)}")

        return True
