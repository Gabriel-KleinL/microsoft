"""
Serviço de Gerenciamento de Projetos
Gerencia projetos e conversas armazenados em arquivos JSON
"""
import json
import os
from typing import List, Dict, Optional
from datetime import datetime


class ProjectsService:
    """Serviço para gerenciar projetos e conversas dos usuários"""

    def __init__(self, data_dir: str = 'backend/data'):
        """
        Inicializa o serviço de projetos

        Args:
            data_dir: Diretório onde os dados serão armazenados
        """
        self.data_dir = data_dir
        self.projects_dir = os.path.join(data_dir, 'projects')

        # Cria diretórios se não existirem
        os.makedirs(self.projects_dir, exist_ok=True)

        print(f"✓ ProjectsService inicializado - Diretório: {self.projects_dir}")

    def _get_user_file(self, user_email: str) -> str:
        """Retorna o caminho do arquivo de projetos do usuário"""
        # Sanitiza o email para usar como nome de arquivo
        safe_email = user_email.replace('@', '_at_').replace('.', '_')
        return os.path.join(self.projects_dir, f'{safe_email}.json')

    def get_all_projects(self, user_email: str) -> List[Dict]:
        """
        Retorna todos os projetos de um usuário

        Args:
            user_email: Email do usuário

        Returns:
            Lista de projetos
        """
        try:
            user_file = self._get_user_file(user_email)

            if not os.path.exists(user_file):
                return []

            with open(user_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
                return data.get('projects', [])

        except Exception as e:
            print(f"❌ Erro ao ler projetos de {user_email}: {e}")
            return []

    def save_projects(self, user_email: str, projects: List[Dict]) -> bool:
        """
        Salva todos os projetos de um usuário

        Args:
            user_email: Email do usuário
            projects: Lista de projetos

        Returns:
            True se salvou com sucesso
        """
        try:
            user_file = self._get_user_file(user_email)

            data = {
                'user_email': user_email,
                'updated_at': datetime.utcnow().isoformat(),
                'projects': projects
            }

            with open(user_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)

            return True

        except Exception as e:
            print(f"❌ Erro ao salvar projetos de {user_email}: {e}")
            return False

    def create_project(self, user_email: str, project_data: Dict) -> Optional[Dict]:
        """
        Cria um novo projeto

        Args:
            user_email: Email do usuário
            project_data: Dados do projeto

        Returns:
            Projeto criado ou None se falhar
        """
        try:
            projects = self.get_all_projects(user_email)

            # Adiciona metadata
            project_data['createdAt'] = datetime.utcnow().isoformat()
            project_data['updatedAt'] = datetime.utcnow().isoformat()

            # Garante que tem a estrutura de conversas
            if 'conversations' not in project_data:
                project_data['conversations'] = [{
                    'id': f"{project_data['id']}_conv1",
                    'title': 'Conversa Principal',
                    'messages': [],
                    'createdAt': datetime.utcnow().isoformat(),
                    'updatedAt': datetime.utcnow().isoformat()
                }]
                project_data['activeConversationId'] = project_data['conversations'][0]['id']

            projects.insert(0, project_data)

            if self.save_projects(user_email, projects):
                return project_data

            return None

        except Exception as e:
            print(f"❌ Erro ao criar projeto: {e}")
            return None

    def update_project(self, user_email: str, project_id: str, updates: Dict) -> Optional[Dict]:
        """
        Atualiza um projeto existente

        Args:
            user_email: Email do usuário
            project_id: ID do projeto
            updates: Campos a atualizar

        Returns:
            Projeto atualizado ou None se não encontrar
        """
        try:
            projects = self.get_all_projects(user_email)

            for i, project in enumerate(projects):
                if project.get('id') == project_id:
                    # Atualiza campos
                    projects[i].update(updates)
                    projects[i]['updatedAt'] = datetime.utcnow().isoformat()

                    if self.save_projects(user_email, projects):
                        return projects[i]
                    break

            return None

        except Exception as e:
            print(f"❌ Erro ao atualizar projeto: {e}")
            return None

    def delete_project(self, user_email: str, project_id: str) -> bool:
        """
        Deleta um projeto

        Args:
            user_email: Email do usuário
            project_id: ID do projeto

        Returns:
            True se deletou com sucesso
        """
        try:
            projects = self.get_all_projects(user_email)
            projects = [p for p in projects if p.get('id') != project_id]
            return self.save_projects(user_email, projects)

        except Exception as e:
            print(f"❌ Erro ao deletar projeto: {e}")
            return False

    def add_conversation(self, user_email: str, project_id: str, conversation: Dict) -> Optional[Dict]:
        """
        Adiciona uma nova conversa a um projeto

        Args:
            user_email: Email do usuário
            project_id: ID do projeto
            conversation: Dados da conversa

        Returns:
            Projeto atualizado ou None se falhar
        """
        try:
            projects = self.get_all_projects(user_email)

            for i, project in enumerate(projects):
                if project.get('id') == project_id:
                    if 'conversations' not in project:
                        project['conversations'] = []

                    conversation['createdAt'] = datetime.utcnow().isoformat()
                    conversation['updatedAt'] = datetime.utcnow().isoformat()

                    project['conversations'].append(conversation)
                    project['activeConversationId'] = conversation['id']
                    project['updatedAt'] = datetime.utcnow().isoformat()

                    projects[i] = project

                    if self.save_projects(user_email, projects):
                        return project
                    break

            return None

        except Exception as e:
            print(f"❌ Erro ao adicionar conversa: {e}")
            return None

    def update_conversation(self, user_email: str, project_id: str, conversation_id: str, updates: Dict) -> Optional[Dict]:
        """
        Atualiza uma conversa de um projeto

        Args:
            user_email: Email do usuário
            project_id: ID do projeto
            conversation_id: ID da conversa
            updates: Campos a atualizar

        Returns:
            Projeto atualizado ou None se falhar
        """
        try:
            projects = self.get_all_projects(user_email)

            for i, project in enumerate(projects):
                if project.get('id') == project_id:
                    if 'conversations' not in project:
                        return None

                    for j, conv in enumerate(project['conversations']):
                        if conv.get('id') == conversation_id:
                            project['conversations'][j].update(updates)
                            project['conversations'][j]['updatedAt'] = datetime.utcnow().isoformat()
                            project['updatedAt'] = datetime.utcnow().isoformat()

                            projects[i] = project

                            if self.save_projects(user_email, projects):
                                return project
                            break
                    break

            return None

        except Exception as e:
            print(f"❌ Erro ao atualizar conversa: {e}")
            return None
