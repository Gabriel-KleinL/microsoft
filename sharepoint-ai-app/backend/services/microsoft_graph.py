"""
Serviço para interação com Microsoft Graph API e SharePoint
"""
import requests
import msal
from typing import Dict, List, Optional
import io


class MicrosoftGraphService:
    """Serviço para autenticação e operações com Microsoft Graph API"""

    def __init__(self, client_id: str, client_secret: str, tenant_id: str):
        self.client_id = client_id
        self.client_secret = client_secret
        self.tenant_id = tenant_id
        self.authority = f"https://login.microsoftonline.com/{tenant_id}"
        self.graph_endpoint = "https://graph.microsoft.com/v1.0"

    def get_authorization_url(self, redirect_uri: str, scopes: List[str]) -> str:
        """
        Gera URL para autorização OAuth 2.0

        Args:
            redirect_uri: URL de redirecionamento após autenticação
            scopes: Lista de permissões solicitadas

        Returns:
            URL de autorização
        """
        app = msal.ConfidentialClientApplication(
            self.client_id,
            authority=self.authority,
            client_credential=self.client_secret
        )

        auth_url = app.get_authorization_request_url(
            scopes=scopes,
            redirect_uri=redirect_uri
        )

        return auth_url

    def acquire_token_by_auth_code(
        self,
        auth_code: str,
        redirect_uri: str,
        scopes: List[str]
    ) -> Optional[Dict]:
        """
        Obtém token de acesso usando código de autorização

        Args:
            auth_code: Código de autorização recebido do OAuth
            redirect_uri: URL de redirecionamento configurada
            scopes: Lista de permissões

        Returns:
            Dicionário com token de acesso e informações
        """
        app = msal.ConfidentialClientApplication(
            self.client_id,
            authority=self.authority,
            client_credential=self.client_secret
        )

        result = app.acquire_token_by_authorization_code(
            code=auth_code,
            scopes=scopes,
            redirect_uri=redirect_uri
        )

        return result

    def get_user_info(self, access_token: str) -> Dict:
        """
        Obtém informações do usuário autenticado

        Args:
            access_token: Token de acesso

        Returns:
            Dicionário com informações do usuário
        """
        headers = {'Authorization': f'Bearer {access_token}'}
        response = requests.get(
            f"{self.graph_endpoint}/me",
            headers=headers
        )
        response.raise_for_status()
        return response.json()

    def get_sharepoint_site_id(self, access_token: str, site_url: str) -> str:
        """
        Obtém o ID do site SharePoint a partir da URL

        Args:
            access_token: Token de acesso
            site_url: URL do site SharePoint

        Returns:
            ID do site SharePoint
        """
        # Extrai hostname e path da URL
        # Exemplo: https://tenant.sharepoint.com/sites/SiteName
        parts = site_url.replace('https://', '').split('/')
        hostname = parts[0]
        site_path = '/' + '/'.join(parts[1:]) if len(parts) > 1 else ''

        headers = {'Authorization': f'Bearer {access_token}'}
        url = f"{self.graph_endpoint}/sites/{hostname}:{site_path}"

        response = requests.get(url, headers=headers)
        response.raise_for_status()
        return response.json()['id']

    def list_documents(
        self,
        access_token: str,
        site_id: str,
        drive_name: str = None
    ) -> List[Dict]:
        """
        Lista documentos do SharePoint

        Args:
            access_token: Token de acesso
            site_id: ID do site SharePoint
            drive_name: Nome da biblioteca de documentos (opcional)

        Returns:
            Lista de documentos
        """
        headers = {'Authorization': f'Bearer {access_token}'}

        # Obtém drives (bibliotecas de documentos) do site
        drives_url = f"{self.graph_endpoint}/sites/{site_id}/drives"
        drives_response = requests.get(drives_url, headers=headers)
        drives_response.raise_for_status()
        drives = drives_response.json().get('value', [])

        all_documents = []

        for drive in drives:
            # Lista itens de cada drive
            items_url = f"{self.graph_endpoint}/drives/{drive['id']}/root/children"

            try:
                items_response = requests.get(items_url, headers=headers)
                items_response.raise_for_status()
                items = items_response.json().get('value', [])

                # Filtra apenas arquivos (não pastas) e formatos suportados
                supported_extensions = ['.pdf', '.docx', '.xlsx', '.pptx', '.txt']

                for item in items:
                    if 'file' in item:  # É um arquivo, não pasta
                        file_name = item.get('name', '')
                        if any(file_name.lower().endswith(ext) for ext in supported_extensions):
                            all_documents.append({
                                'id': item['id'],
                                'name': file_name,
                                'size': item.get('size', 0),
                                'webUrl': item.get('webUrl', ''),
                                'downloadUrl': item.get('@microsoft.graph.downloadUrl', ''),
                                'driveId': drive['id'],
                                'lastModified': item.get('lastModifiedDateTime', ''),
                                'type': self._get_file_type(file_name)
                            })
            except requests.exceptions.HTTPError as e:
                print(f"Erro ao acessar drive {drive.get('name', 'unknown')}: {e}")
                continue

        return all_documents

    def download_file_content(self, access_token: str, drive_id: str, file_id: str) -> bytes:
        """
        Baixa conteúdo de um arquivo do SharePoint

        Args:
            access_token: Token de acesso
            drive_id: ID do drive
            file_id: ID do arquivo

        Returns:
            Conteúdo do arquivo em bytes
        """
        headers = {'Authorization': f'Bearer {access_token}'}
        url = f"{self.graph_endpoint}/drives/{drive_id}/items/{file_id}/content"

        response = requests.get(url, headers=headers)
        response.raise_for_status()

        return response.content

    @staticmethod
    def _get_file_type(filename: str) -> str:
        """
        Determina o tipo de arquivo baseado na extensão

        Args:
            filename: Nome do arquivo

        Returns:
            Tipo do arquivo
        """
        filename_lower = filename.lower()
        if filename_lower.endswith('.pdf'):
            return 'pdf'
        elif filename_lower.endswith('.docx'):
            return 'word'
        elif filename_lower.endswith('.xlsx'):
            return 'excel'
        elif filename_lower.endswith('.pptx'):
            return 'powerpoint'
        elif filename_lower.endswith('.txt'):
            return 'text'
        else:
            return 'unknown'
