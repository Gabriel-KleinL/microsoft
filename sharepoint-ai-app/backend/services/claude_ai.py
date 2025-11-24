"""
Serviço para interação com Claude API (Anthropic)
"""
import anthropic
import base64
from typing import Dict, List
import PyPDF2
import io
from docx import Document
from pptx import Presentation
import openpyxl


class ClaudeAIService:
    """Serviço para processamento de documentos com Claude AI"""

    def __init__(self, api_key: str, model: str = "claude-3-5-sonnet-20241022"):
        self.client = anthropic.Anthropic(api_key=api_key)
        self.model = model

    def extract_text_from_pdf(self, file_content: bytes) -> str:
        """
        Extrai texto de arquivo PDF

        Args:
            file_content: Conteúdo do arquivo PDF em bytes

        Returns:
            Texto extraído do PDF
        """
        try:
            pdf_file = io.BytesIO(file_content)
            pdf_reader = PyPDF2.PdfReader(pdf_file)

            text = []
            for page in pdf_reader.pages:
                text.append(page.extract_text())

            return '\n'.join(text)
        except Exception as e:
            return f"Erro ao extrair texto do PDF: {str(e)}"

    def extract_text_from_docx(self, file_content: bytes) -> str:
        """
        Extrai texto de arquivo Word (.docx)

        Args:
            file_content: Conteúdo do arquivo DOCX em bytes

        Returns:
            Texto extraído do documento
        """
        try:
            doc_file = io.BytesIO(file_content)
            doc = Document(doc_file)

            text = []
            for paragraph in doc.paragraphs:
                text.append(paragraph.text)

            # Também extrai texto de tabelas
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        text.append(cell.text)

            return '\n'.join(text)
        except Exception as e:
            return f"Erro ao extrair texto do Word: {str(e)}"

    def extract_text_from_xlsx(self, file_content: bytes) -> str:
        """
        Extrai texto de arquivo Excel (.xlsx)

        Args:
            file_content: Conteúdo do arquivo XLSX em bytes

        Returns:
            Texto extraído da planilha
        """
        try:
            excel_file = io.BytesIO(file_content)
            workbook = openpyxl.load_workbook(excel_file, data_only=True)

            text = []
            for sheet_name in workbook.sheetnames:
                sheet = workbook[sheet_name]
                text.append(f"\n=== Planilha: {sheet_name} ===\n")

                for row in sheet.iter_rows(values_only=True):
                    row_text = '\t'.join([str(cell) if cell is not None else '' for cell in row])
                    if row_text.strip():
                        text.append(row_text)

            return '\n'.join(text)
        except Exception as e:
            return f"Erro ao extrair texto do Excel: {str(e)}"

    def extract_text_from_pptx(self, file_content: bytes) -> str:
        """
        Extrai texto de arquivo PowerPoint (.pptx)

        Args:
            file_content: Conteúdo do arquivo PPTX em bytes

        Returns:
            Texto extraído da apresentação
        """
        try:
            ppt_file = io.BytesIO(file_content)
            presentation = Presentation(ppt_file)

            text = []
            for i, slide in enumerate(presentation.slides, 1):
                text.append(f"\n=== Slide {i} ===\n")

                for shape in slide.shapes:
                    if hasattr(shape, "text"):
                        text.append(shape.text)

            return '\n'.join(text)
        except Exception as e:
            return f"Erro ao extrair texto do PowerPoint: {str(e)}"

    def extract_text_from_file(self, file_content: bytes, file_type: str) -> str:
        """
        Extrai texto de arquivo baseado no tipo

        Args:
            file_content: Conteúdo do arquivo em bytes
            file_type: Tipo do arquivo (pdf, word, excel, powerpoint, text)

        Returns:
            Texto extraído
        """
        if file_type == 'pdf':
            return self.extract_text_from_pdf(file_content)
        elif file_type == 'word':
            return self.extract_text_from_docx(file_content)
        elif file_type == 'excel':
            return self.extract_text_from_xlsx(file_content)
        elif file_type == 'powerpoint':
            return self.extract_text_from_pptx(file_content)
        elif file_type == 'text':
            return file_content.decode('utf-8', errors='ignore')
        else:
            return "Tipo de arquivo não suportado"

    def summarize_document(
        self,
        file_content: bytes,
        file_type: str,
        file_name: str,
        language: str = "pt-BR"
    ) -> Dict:
        """
        Resume um documento usando Claude AI

        Args:
            file_content: Conteúdo do arquivo em bytes
            file_type: Tipo do arquivo
            file_name: Nome do arquivo
            language: Idioma do resumo (padrão: pt-BR)

        Returns:
            Dicionário com resumo e informações
        """
        try:
            # Extrai texto do documento
            extracted_text = self.extract_text_from_file(file_content, file_type)

            if extracted_text.startswith("Erro"):
                return {
                    'success': False,
                    'error': extracted_text,
                    'file_name': file_name
                }

            # Limita o texto se for muito grande (Claude tem limite de tokens)
            max_chars = 100000  # Aproximadamente 25k tokens
            if len(extracted_text) > max_chars:
                extracted_text = extracted_text[:max_chars] + "\n\n[... documento truncado ...]"

            # Prompt para resumo em português
            prompt = f"""Você é um assistente especializado em análise de documentos.

Analise o seguinte documento e forneça:

1. **Resumo Executivo**: Um resumo conciso do conteúdo principal (2-3 parágrafos)
2. **Pontos Principais**: Liste os 5-7 pontos mais importantes do documento
3. **Informações Chave**: Dados, números ou fatos relevantes mencionados
4. **Categoria**: Qual o tipo/categoria deste documento (ex: relatório, apresentação, planilha, etc.)

Documento: {file_name}

Conteúdo:
{extracted_text}

Por favor, forneça a análise em português do Brasil, de forma clara e profissional."""

            # Chama Claude API
            message = self.client.messages.create(
                model=self.model,
                max_tokens=2000,
                messages=[
                    {"role": "user", "content": prompt}
                ]
            )

            summary = message.content[0].text

            return {
                'success': True,
                'file_name': file_name,
                'file_type': file_type,
                'summary': summary,
                'char_count': len(extracted_text),
                'model_used': self.model
            }

        except Exception as e:
            return {
                'success': False,
                'error': f"Erro ao processar documento: {str(e)}",
                'file_name': file_name
            }

    def summarize_multiple_documents(
        self,
        documents: List[Dict],
        language: str = "pt-BR"
    ) -> List[Dict]:
        """
        Resume múltiplos documentos

        Args:
            documents: Lista de documentos com 'content', 'type' e 'name'
            language: Idioma do resumo

        Returns:
            Lista de resumos
        """
        summaries = []

        for doc in documents:
            summary = self.summarize_document(
                file_content=doc['content'],
                file_type=doc['type'],
                file_name=doc['name'],
                language=language
            )
            summaries.append(summary)

        return summaries

    def generate_combined_summary(self, individual_summaries: List[str], file_names: List[str]) -> str:
        """
        Gera um resumo consolidado de múltiplos documentos

        Args:
            individual_summaries: Lista de resumos individuais
            file_names: Lista de nomes dos arquivos

        Returns:
            Resumo consolidado
        """
        try:
            # Cria texto com todos os resumos
            combined_text = ""
            for name, summary in zip(file_names, individual_summaries):
                combined_text += f"\n\n### {name}\n{summary}\n"

            prompt = f"""Você recebeu resumos de múltiplos documentos.

Crie um resumo executivo consolidado que:
1. Identifique temas comuns entre os documentos
2. Destaque as informações mais importantes de todos os documentos
3. Organize as informações de forma lógica e coerente
4. Forneça insights sobre como os documentos se relacionam (se aplicável)

Resumos dos documentos:
{combined_text}

Forneça o resumo consolidado em português do Brasil, de forma profissional."""

            message = self.client.messages.create(
                model=self.model,
                max_tokens=2000,
                messages=[
                    {"role": "user", "content": prompt}
                ]
            )

            return message.content[0].text

        except Exception as e:
            return f"Erro ao gerar resumo consolidado: {str(e)}"
