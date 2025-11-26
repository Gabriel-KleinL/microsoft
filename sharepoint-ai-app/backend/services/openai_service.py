"""
Serviço OpenAI - Integração com ChatGPT
"""
from openai import OpenAI
from typing import List, Dict, Optional
import json
from .claude_ai import ClaudeAIService


class OpenAIService:
    """
    Serviço para integração com OpenAI ChatGPT API
    """

    def __init__(self, api_key: str, model: str = "gpt-4"):
        """
        Inicializa o serviço OpenAI

        Args:
            api_key: Chave de API da OpenAI
            model: Modelo a ser usado (gpt-4, gpt-4-turbo, gpt-3.5-turbo)
        """
        self.client = OpenAI(api_key=api_key)
        self.model = model

        # Cria uma instância auxiliar do ClaudeAIService apenas para extração de texto
        # (a extração de PDF/Word/Excel é independente da IA usada)
        self._text_extractor = ClaudeAIService(api_key="dummy", model="dummy")

        print(f"✓ OpenAI Service inicializado - Modelo: {model}")

    def chat(
        self,
        messages: List[Dict[str, str]],
        max_tokens: int = 2000,
        temperature: float = 0.7
    ) -> str:
        """
        Envia mensagens para o ChatGPT e retorna a resposta

        Args:
            messages: Lista de mensagens no formato [{"role": "user", "content": "..."}]
            max_tokens: Número máximo de tokens na resposta
            temperature: Criatividade da resposta (0.0 a 1.0)

        Returns:
            Resposta do ChatGPT
        """
        try:
            response = self.client.chat.completions.create(
                model=self.model,
                messages=messages,
                max_tokens=max_tokens,
                temperature=temperature
            )

            return response.choices[0].message.content

        except Exception as e:
            print(f"❌ Erro ao chamar OpenAI API: {e}")
            raise

    def summarize_document(
        self,
        file_name: str,
        file_content: bytes,
        file_type: str
    ) -> Dict:
        """
        Resume um documento usando ChatGPT

        Args:
            file_name: Nome do arquivo
            file_content: Conteúdo do arquivo em bytes
            file_type: Tipo do arquivo (pdf, word, excel, etc)

        Returns:
            Dicionário com resumo e informações
        """
        try:
            # Extrai texto do arquivo usando o método apropriado para cada tipo
            if isinstance(file_content, bytes):
                content_str = self.extract_text_from_file(file_content, file_type, file_name)

                # Se a extração retornou erro, trata como falha
                if content_str.startswith("Erro"):
                    return {
                        'success': False,
                        'error': content_str,
                        'file_name': file_name
                    }
            else:
                content_str = str(file_content)

            system_prompt = """Você é um assistente especializado em analisar e resumir documentos corporativos.
Sua tarefa é criar resumos concisos, informativos e bem estruturados.

Diretrizes:
- Identifique os pontos principais e informações mais relevantes
- Use formatação markdown para melhor legibilidade
- Destaque dados importantes, datas, valores e decisões
- Seja objetivo e direto
- Use bullets quando apropriado
"""

            # Limita o conteúdo para evitar exceder o limite de tokens
            # GPT-4 tem limite de 8192 tokens, precisamos deixar espaço para:
            # - System prompt (~200 tokens)
            # - User prompt structure (~150 tokens)
            # - Response (1000 tokens)
            # Isso deixa ~6800 tokens para o conteúdo, que equivale a ~3400 caracteres
            max_content_chars = 3000
            if len(content_str) > max_content_chars:
                content_preview = content_str[:max_content_chars] + f"\n\n[... documento truncado - {len(content_str)} caracteres no total ...]"
            else:
                content_preview = content_str

            user_prompt = f"""Analise e resuma o seguinte documento:

**Arquivo:** {file_name}
**Tipo:** {file_type}

**Conteúdo:**
{content_preview}

Forneça um resumo estruturado destacando:
1. Assunto principal
2. Pontos-chave
3. Informações importantes (datas, valores, pessoas, etc)
4. Conclusões ou ações necessárias (se houver)
"""

            messages = [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt}
            ]

            summary_text = self.chat(messages, max_tokens=800)

            return {
                'success': True,
                'file_name': file_name,
                'file_type': file_type,
                'summary': summary_text,
                'char_count': len(content_str),
                'model_used': self.model
            }

        except Exception as e:
            return {
                'success': False,
                'error': f"Erro ao processar documento: {str(e)}",
                'file_name': file_name
            }

    def consolidate_summaries(
        self,
        summaries: List[Dict],
        document_count: int
    ) -> str:
        """
        Consolida múltiplos resumos em um resumo geral

        Args:
            summaries: Lista de resumos individuais
            document_count: Número total de documentos

        Returns:
            Resumo consolidado
        """
        summaries_text = "\n\n".join([
            f"**{s['file_name']}:**\n{s['summary']}"
            for s in summaries if s.get('success')
        ])

        system_prompt = """Você é um assistente especializado em consolidar informações de múltiplos documentos.
Crie um resumo executivo que integre as informações de todos os documentos."""

        user_prompt = f"""Analise os seguintes {document_count} resumos de documentos e crie um resumo consolidado:

{summaries_text}

Crie um resumo executivo que:
1. Identifique temas comuns entre os documentos
2. Destaque as informações mais importantes
3. Organize por tópicos quando possível
4. Seja conciso mas informativo
"""

        messages = [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt}
        ]

        return self.chat(messages, max_tokens=1500)

    def answer_question(
        self,
        question: str,
        documents_context: str,
        conversation_history: Optional[List[Dict]] = None
    ) -> Dict:
        """
        Responde perguntas baseadas em documentos do SharePoint

        Args:
            question: Pergunta do usuário
            documents_context: Contexto dos documentos relevantes
            conversation_history: Histórico da conversa (opcional)

        Returns:
            Dict com a resposta e metadados
        """
        system_prompt = f"""Você é Sofia, uma assistente de IA especializada em ajudar usuários a encontrar informações em documentos do SharePoint.

**Contexto dos Documentos:**
{documents_context}

**Instruções:**
- Responda com base APENAS nas informações dos documentos fornecidos
- Se a informação não estiver nos documentos, diga claramente que não encontrou
- Cite os nomes dos documentos quando relevante
- Seja precisa, útil e amigável
- Use formatação markdown para melhor legibilidade
- Se houver múltiplas respostas possíveis, liste todas
"""

        messages = [
            {"role": "system", "content": system_prompt}
        ]

        # Adiciona histórico se houver
        if conversation_history:
            messages.extend(conversation_history[-5:])  # Últimas 5 mensagens

        # Adiciona pergunta atual
        messages.append({
            "role": "user",
            "content": question
        })

        response = self.chat(messages, max_tokens=2000, temperature=0.7)

        return {
            "response": response,
            "model": self.model,
            "provider": "openai"
        }

    def extract_text_from_file(self, file_content: bytes, file_type: str, file_name: str = "") -> str:
        """
        Extrai texto de arquivo (reutiliza a lógica do ClaudeAI Service)

        Args:
            file_content: Conteúdo do arquivo em bytes
            file_type: Tipo do arquivo
            file_name: Nome do arquivo (opcional)

        Returns:
            Texto extraído
        """
        return self._text_extractor.extract_text_from_file(file_content, file_type, file_name)

    def generate_combined_summary(self, individual_summaries: list, file_names: list) -> str:
        """
        Gera um resumo consolidado de múltiplos documentos

        Args:
            individual_summaries: Lista de resumos individuais
            file_names: Lista de nomes dos arquivos

        Returns:
            Resumo consolidado
        """
        summaries_text = "\n\n".join([
            f"**{name}:**\n{summary}"
            for name, summary in zip(file_names, individual_summaries)
        ])

        system_prompt = """Você é um assistente especializado em consolidar informações de múltiplos documentos.
Crie um resumo executivo que integre as informações de todos os documentos."""

        user_prompt = f"""Analise os seguintes resumos de documentos e crie um resumo consolidado:

{summaries_text}

Crie um resumo executivo que:
1. Identifique temas comuns entre os documentos
2. Destaque as informações mais importantes
3. Organize por tópicos quando possível
4. Seja conciso mas informativo
"""

        messages = [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt}
        ]

        return self.chat(messages, max_tokens=1500)
