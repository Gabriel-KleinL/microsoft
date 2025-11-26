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

    def _get_all_items_paginated(self, url: str, headers: Dict, max_items: int = 5000) -> List[Dict]:
        """
        Busca todos os itens de uma URL com paginação

        Args:
            url: URL inicial
            headers: Headers HTTP
            max_items: Máximo de itens a buscar (proteção contra loops infinitos)

        Returns:
            Lista completa de itens
        """
        all_items = []
        current_url = url
        page_count = 0

        while current_url and len(all_items) < max_items:
            page_count += 1
            response = requests.get(current_url, headers=headers)
            response.raise_for_status()
            data = response.json()

            items = data.get('value', [])
            all_items.extend(items)

            # Verifica se há próxima página
            current_url = data.get('@odata.nextLink', None)

            if current_url:
                print(f"  📄 Página {page_count}: {len(items)} itens (total: {len(all_items)})")

        if len(all_items) > 0:
            print(f"  ✓ Total: {len(all_items)} itens em {page_count} página(s)")

        return all_items

    def list_documents(
        self,
        access_token: str,
        site_id: str,
        drive_name: str = None,
        folder_path: str = None
    ) -> List[Dict]:
        """
        Lista documentos e pastas do SharePoint com paginação completa e hierarquia

        Args:
            access_token: Token de acesso
            site_id: ID do site SharePoint
            drive_name: Nome da biblioteca de documentos (opcional)
            folder_path: Caminho da pasta no formato 'drive:{drive_id}' ou 'drive:{drive_id}/path'

        Returns:
            Lista de documentos e pastas
        """
        headers = {'Authorization': f'Bearer {access_token}'}

        # Se folder_path é None (raiz), retorna os drives como pastas principais
        if not folder_path:
            return self._list_drives_as_folders(site_id, headers)

        # Caso contrário, navega dentro de um drive específico
        return self._list_drive_contents(folder_path, headers)

    def _list_drives_as_folders(self, site_id: str, headers: Dict) -> List[Dict]:
        """
        Lista os drives (bibliotecas) como pastas principais

        Args:
            site_id: ID do site SharePoint
            headers: Headers HTTP

        Returns:
            Lista de drives formatados como pastas
        """
        drives_url = f"{self.graph_endpoint}/sites/{site_id}/drives"
        print(f"🔍 Listando bibliotecas principais (raiz)...")
        drives = self._get_all_items_paginated(drives_url, headers)

        all_items = []

        for drive in drives:
            drive_name = drive.get('name', 'Desconhecido')

            # Cria uma "pasta" para cada drive
            all_items.append({
                'id': drive['id'],
                'name': drive_name,
                'size': 0,
                'webUrl': drive.get('webUrl', ''),
                'driveId': drive['id'],
                'driveName': drive_name,
                'lastModified': drive.get('lastModifiedDateTime', ''),
                'lastModifiedBy': drive.get('lastModifiedBy', {}).get('user', {}).get('displayName', ''),
                'type': 'folder',
                'isFolder': True,
                'isDrive': True,  # Marca como drive/biblioteca
                'childCount': 0,
                'parentPath': '',
                'folderPath': f"drive:{drive['id']}"  # Caminho para navegação
            })

        # Ordena alfabeticamente
        all_items.sort(key=lambda x: x.get('name', '').lower())

        print(f"✓ {len(all_items)} bibliotecas encontradas")

        return all_items

    def _list_drive_contents(self, folder_path: str, headers: Dict) -> List[Dict]:
        """
        Lista conteúdos dentro de um drive específico

        Args:
            folder_path: Caminho no formato 'drive:{drive_id}' ou 'drive:{drive_id}/path'
            headers: Headers HTTP

        Returns:
            Lista de itens (pastas e arquivos)
        """
        # Parse do folder_path
        # Formato: drive:{drive_id} ou drive:{drive_id}/caminho/subpasta
        if not folder_path.startswith('drive:'):
            raise ValueError(f"Formato de caminho inválido: {folder_path}")

        # Remove o prefixo 'drive:'
        path_parts = folder_path[6:].split('/', 1)
        drive_id = path_parts[0]
        internal_path = path_parts[1] if len(path_parts) > 1 else None

        print(f"📂 Navegando: Drive ID={drive_id}, Caminho interno={internal_path or 'raiz'}")

        # Monta a URL baseada se há caminho interno
        if internal_path:
            items_url = f"{self.graph_endpoint}/drives/{drive_id}/root:/{internal_path}:/children"
        else:
            items_url = f"{self.graph_endpoint}/drives/{drive_id}/root/children"

        # Busca os itens
        items = self._get_all_items_paginated(items_url, headers)

        all_items = []
        total_files_found = 0
        total_folders_found = 0

        for item in items:
            # Verifica se é pasta
            if 'folder' in item:
                # Monta o novo caminho para navegação
                new_path = f"drive:{drive_id}"
                if internal_path:
                    new_path += f"/{internal_path}/{item.get('name', '')}"
                else:
                    new_path += f"/{item.get('name', '')}"

                all_items.append({
                    'id': item['id'],
                    'name': item.get('name', ''),
                    'size': 0,
                    'webUrl': item.get('webUrl', ''),
                    'driveId': drive_id,
                    'driveName': '',  # Preenchido depois se necessário
                    'lastModified': item.get('lastModifiedDateTime', ''),
                    'lastModifiedBy': item.get('lastModifiedBy', {}).get('user', {}).get('displayName', ''),
                    'type': 'folder',
                    'isFolder': True,
                    'isDrive': False,
                    'childCount': item.get('folder', {}).get('childCount', 0),
                    'parentPath': folder_path,
                    'folderPath': new_path
                })
                total_folders_found += 1

            # Verifica se é arquivo (aceita TODOS os arquivos)
            elif 'file' in item:
                file_name = item.get('name', '')
                all_items.append({
                    'id': item['id'],
                    'name': file_name,
                    'size': item.get('size', 0),
                    'webUrl': item.get('webUrl', ''),
                    'downloadUrl': item.get('@microsoft.graph.downloadUrl', ''),
                    'driveId': drive_id,
                    'driveName': '',
                    'lastModified': item.get('lastModifiedDateTime', ''),
                    'lastModifiedBy': item.get('lastModifiedBy', {}).get('user', {}).get('displayName', ''),
                    'type': self._get_file_type(file_name),
                    'isFolder': False,
                    'parentPath': folder_path
                })
                total_files_found += 1

        # Ordena: pastas primeiro, depois arquivos (ambos alfabeticamente)
        all_items.sort(key=lambda x: (not x.get('isFolder', False), x.get('name', '').lower()))

        print(f"✓ Encontrados: {total_folders_found} pastas, {total_files_found} arquivos")

        return all_items

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

        # Documentos PDF
        if filename_lower.endswith('.pdf'):
            return 'pdf'

        # Documentos Word
        elif filename_lower.endswith(('.docx', '.doc')):
            return 'word'

        # Planilhas Excel
        elif filename_lower.endswith(('.xlsx', '.xls', '.xlsm', '.csv')):
            return 'excel'

        # Apresentações PowerPoint
        elif filename_lower.endswith(('.pptx', '.ppt')):
            return 'powerpoint'

        # Arquivos de texto
        elif filename_lower.endswith(('.txt', '.md', '.log', '.json', '.xml', '.html', '.css', '.js', '.py', '.java', '.cpp', '.c', '.h')):
            return 'text'

        # Imagens
        elif filename_lower.endswith(('.jpg', '.jpeg', '.png', '.gif', '.bmp', '.svg', '.ico', '.webp')):
            return 'image'

        # Vídeos
        elif filename_lower.endswith(('.mp4', '.avi', '.mov', '.wmv', '.flv', '.mkv', '.webm')):
            return 'video'

        # Áudio
        elif filename_lower.endswith(('.mp3', '.wav', '.flac', '.aac', '.ogg', '.wma', '.m4a')):
            return 'audio'

        # Arquivos comprimidos
        elif filename_lower.endswith(('.zip', '.rar', '.7z', '.tar', '.gz', '.bz2')):
            return 'archive'

        # Executáveis/Aplicativos
        elif filename_lower.endswith(('.exe', '.msi', '.app', '.apk', '.deb', '.rpm')):
            return 'executable'

        # Outros tipos conhecidos
        elif filename_lower.endswith('.dwg'):
            return 'cad'
        elif filename_lower.endswith(('.psd', '.ai', '.sketch')):
            return 'design'

        # Tipo desconhecido
        else:
            # Tenta extrair a extensão para mostrar
            parts = filename.rsplit('.', 1)
            if len(parts) > 1:
                return parts[1].upper()
            return 'file'

    def search_all_documents(
        self,
        access_token: str,
        site_id: str,
        query: str,
        max_results: int = 200
    ) -> List[Dict]:
        """
        Busca documentos em todo o SharePoint usando Microsoft Graph Search API

        Args:
            access_token: Token de acesso
            site_id: ID do site SharePoint
            query: Termo de busca
            max_results: Máximo de resultados (padrão: 200)

        Returns:
            Lista de documentos encontrados
        """
        headers = {
            'Authorization': f'Bearer {access_token}',
            'Content-Type': 'application/json'
        }

        # Usa a Microsoft Graph Search API
        search_url = f"{self.graph_endpoint}/search/query"

        search_body = {
            "requests": [
                {
                    "entityTypes": ["driveItem"],
                    "query": {
                        "queryString": f"{query}"
                    },
                    "from": 0,
                    "size": max_results,
                    "fields": [
                        "name",
                        "id",
                        "size",
                        "webUrl",
                        "lastModifiedDateTime",
                        "createdDateTime",
                        "parentReference"
                    ]
                }
            ]
        }

        try:
            print(f"🔍 Buscando '{query}' em todo o SharePoint...")
            response = requests.post(search_url, headers=headers, json=search_body)
            response.raise_for_status()
            data = response.json()

            all_items = []

            # Processa resultados da busca
            if 'value' in data and len(data['value']) > 0:
                hits_container = data['value'][0].get('hitsContainers', [])

                for container in hits_container:
                    hits = container.get('hits', [])

                    for hit in hits:
                        resource = hit.get('resource', {})

                        # Extrai informações do documento
                        file_name = resource.get('name', '')
                        file_id = resource.get('id', '')

                        # Verifica se é um arquivo (não pasta)
                        if file_name and not resource.get('folder'):
                            parent_ref = resource.get('parentReference', {})

                            item = {
                                'id': file_id,
                                'name': file_name,
                                'size': resource.get('size', 0),
                                'webUrl': resource.get('webUrl', ''),
                                'driveId': parent_ref.get('driveId', ''),
                                'driveName': parent_ref.get('name', ''),
                                'lastModified': resource.get('lastModifiedDateTime', ''),
                                'lastModifiedBy': resource.get('lastModifiedBy', {}).get('user', {}).get('displayName', ''),
                                'type': self._get_file_type(file_name),
                                'isFolder': False,
                                'parentPath': resource.get('parentReference', {}).get('path', ''),
                                'searchScore': hit.get('rank', 0)
                            }

                            all_items.append(item)

            print(f"✓ Encontrados {len(all_items)} documentos na busca global")
            return all_items

        except Exception as e:
            print(f"⚠ Erro na busca do Graph Search API: {e}")
            # Fallback: busca recursiva manual
            return self._search_recursively(access_token, site_id, query, max_results)

    def _search_recursively(
        self,
        access_token: str,
        site_id: str,
        query: str,
        max_results: int = 200
    ) -> List[Dict]:
        """
        Busca recursiva manual como fallback (caso Graph Search API falhe)

        Args:
            access_token: Token de acesso
            site_id: ID do site SharePoint
            query: Termo de busca
            max_results: Máximo de resultados

        Returns:
            Lista de documentos encontrados
        """
        headers = {'Authorization': f'Bearer {access_token}'}
        query_lower = query.lower()
        all_matches = []

        print(f"🔄 Usando busca recursiva manual para '{query}'...")

        try:
            # Obtém todos os drives
            drives_url = f"{self.graph_endpoint}/sites/{site_id}/drives"
            drives = self._get_all_items_paginated(drives_url, headers, max_items=100)

            for drive in drives:
                drive_id = drive['id']
                drive_name = drive.get('name', '')

                print(f"  📂 Buscando em: {drive_name}")

                # Busca recursiva neste drive
                matches = self._search_in_drive(
                    access_token,
                    drive_id,
                    drive_name,
                    query_lower,
                    max_results - len(all_matches)
                )

                all_matches.extend(matches)

                # Para se atingir o máximo
                if len(all_matches) >= max_results:
                    break

            print(f"✓ Busca recursiva encontrou {len(all_matches)} documentos")
            return all_matches[:max_results]

        except Exception as e:
            print(f"❌ Erro na busca recursiva: {e}")
            return []

    def _search_in_drive(
        self,
        access_token: str,
        drive_id: str,
        drive_name: str,
        query: str,
        max_results: int
    ) -> List[Dict]:
        """
        Busca recursivamente em um drive específico

        Args:
            access_token: Token de acesso
            drive_id: ID do drive
            drive_name: Nome do drive
            query: Termo de busca (lowercase)
            max_results: Máximo de resultados

        Returns:
            Lista de documentos encontrados
        """
        headers = {'Authorization': f'Bearer {access_token}'}
        matches = []

        def search_folder(folder_path: str = None):
            """Função recursiva para buscar em pastas"""
            if len(matches) >= max_results:
                return

            # Monta URL
            if folder_path:
                items_url = f"{self.graph_endpoint}/drives/{drive_id}/root:/{folder_path}:/children"
            else:
                items_url = f"{self.graph_endpoint}/drives/{drive_id}/root/children"

            try:
                items = self._get_all_items_paginated(items_url, headers, max_items=1000)

                for item in items:
                    if len(matches) >= max_results:
                        break

                    item_name = item.get('name', '').lower()

                    # Se é arquivo e corresponde à busca
                    if 'file' in item and query in item_name:
                        matches.append({
                            'id': item['id'],
                            'name': item.get('name', ''),
                            'size': item.get('size', 0),
                            'webUrl': item.get('webUrl', ''),
                            'driveId': drive_id,
                            'driveName': drive_name,
                            'lastModified': item.get('lastModifiedDateTime', ''),
                            'lastModifiedBy': item.get('lastModifiedBy', {}).get('user', {}).get('displayName', ''),
                            'type': self._get_file_type(item.get('name', '')),
                            'isFolder': False,
                            'parentPath': folder_path or '',
                            'downloadUrl': item.get('@microsoft.graph.downloadUrl', '')
                        })

                    # Se é pasta, busca recursivamente
                    elif 'folder' in item:
                        folder_name = item.get('name', '')
                        new_path = f"{folder_path}/{folder_name}" if folder_path else folder_name
                        search_folder(new_path)

            except Exception as e:
                # Ignora erros de acesso (pastas sem permissão, etc)
                pass

        # Inicia busca recursiva
        search_folder()
        return matches
