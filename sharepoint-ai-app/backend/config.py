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
        required = [
            'MICROSOFT_CLIENT_ID',
            'MICROSOFT_CLIENT_SECRET',
            'MICROSOFT_TENANT_ID',
            'SHAREPOINT_SITE_URL',
            'CLAUDE_API_KEY'
        ]

        missing = []
        for key in required:
            if not getattr(Config, key):
                missing.append(key)

        if missing:
            raise ValueError(f"Configurações ausentes: {', '.join(missing)}")

        return True
