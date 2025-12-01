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

    def __init__(self, api_key: str, model: str = "claude-sonnet-4-5-20250929"):
        self.client = anthropic.Anthropic(api_key=api_key)
        self.model = model

    def extract_text_from_pdf(self, file_content: bytes) -> str:
        """
        Extrai texto de arquivo PDF usando múltiplos métodos

        Args:
            file_content: Conteúdo do arquivo PDF em bytes

        Returns:
            Texto extraído do PDF
        """
        text_result = ""

        # Método 1: Tenta PyPDF2 (rápido)
        try:
            pdf_file = io.BytesIO(file_content)
            pdf_reader = PyPDF2.PdfReader(pdf_file)

            text = []
            for page in pdf_reader.pages:
                page_text = page.extract_text()
                if page_text:
                    text.append(page_text)

            text_result = '\n'.join(text)

            print(f"📄 PyPDF2: Extraído {len(text_result)} caracteres de {len(pdf_reader.pages)} páginas")

            # Se conseguiu extrair texto suficiente, retorna
            if len(text_result.strip()) > 50:
                print(f"✅ PyPDF2 sucesso!")
                return text_result
        except Exception as e:
            print(f"❌ PyPDF2 falhou: {str(e)}")

        # Método 2: Tenta pdfplumber (mais robusto)
        try:
            import pdfplumber

            pdf_file = io.BytesIO(file_content)
            with pdfplumber.open(pdf_file) as pdf:
                text = []
                for page in pdf.pages:
                    page_text = page.extract_text()
                    if page_text:
                        text.append(page_text)

                text_result = '\n'.join(text)

                print(f"📄 pdfplumber: Extraído {len(text_result)} caracteres de {len(pdf.pages)} páginas")

                # Se conseguiu extrair texto suficiente, retorna
                if len(text_result.strip()) > 50:
                    print(f"✅ pdfplumber sucesso!")
                    return text_result
        except Exception as e:
            print(f"❌ pdfplumber falhou: {str(e)}")

        # Se ambos os métodos falharam ou não extraíram texto suficiente
        if len(text_result.strip()) < 50:
            print("⚠️ PDF com pouco texto extraído. Retornando None para fallback.")
            return None

        return text_result

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
        Extrai texto de arquivo Excel (.xlsx) ou CSV

        Args:
            file_content: Conteúdo do arquivo XLSX/CSV em bytes

        Returns:
            Texto extraído da planilha
        """
        try:
            # Tenta primeiro como CSV (arquivo de texto)
            try:
                csv_text = file_content.decode('utf-8', errors='ignore')
                # Verifica se parece ser um CSV válido (tem linhas com delimitadores)
                if ',' in csv_text or ';' in csv_text or '\t' in csv_text:
                    lines = csv_text.split('\n')
                    if len(lines) > 0:
                        # É um CSV válido
                        import csv
                        import io as csv_io

                        # Detecta o delimitador
                        delimiter = ','
                        if ';' in lines[0]:
                            delimiter = ';'
                        elif '\t' in lines[0]:
                            delimiter = '\t'

                        # Processa o CSV
                        text = ["\n=== Planilha CSV ===\n"]
                        csv_reader = csv.reader(csv_io.StringIO(csv_text), delimiter=delimiter)
                        for row in csv_reader:
                            row_text = '\t'.join([str(cell) for cell in row])
                            if row_text.strip():
                                text.append(row_text)

                        return '\n'.join(text)
            except:
                pass

            # Se não for CSV, tenta como XLSX (arquivo binário)
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
            return f"Erro ao extrair texto do Excel/CSV: {str(e)}"

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

    def analyze_image(self, file_content: bytes, file_name: str) -> str:
        """
        Analisa imagem usando Claude Vision API
        
        Args:
            file_content: Conteúdo da imagem em bytes
            file_name: Nome do arquivo
            
        Returns:
            Descrição da imagem
        """
        try:
            # Detecta tipo de imagem
            import imghdr
            image_type = imghdr.what(None, h=file_content)
            
            if not image_type:
                # Tenta detectar pelo nome do arquivo
                ext = file_name.lower().split('.')[-1]
                image_type = ext if ext in ['jpg', 'jpeg', 'png', 'gif', 'webp'] else 'jpeg'
            
            # Converte para base64
            image_data = base64.standard_b64encode(file_content).decode("utf-8")
            
            # Usa Claude Vision para analisar a imagem
            message = self.client.messages.create(
                model=self.model,
                max_tokens=1024,
                messages=[
                    {
                        "role": "user",
                        "content": [
                            {
                                "type": "image",
                                "source": {
                                    "type": "base64",
                                    "media_type": f"image/{image_type}",
                                    "data": image_data,
                                },
                            },
                            {
                                "type": "text",
                                "text": "Descreva esta imagem em detalhes em português do Brasil. Inclua: o que você vê, elementos principais, cores, texto visível (se houver), contexto e qualquer informação relevante."
                            }
                        ],
                    }
                ],
            )
            
            return f"[ANÁLISE DE IMAGEM]\n{message.content[0].text}"
            
        except Exception as e:
            return f"[IMAGEM] Não foi possível analisar a imagem: {str(e)}"
    
    def analyze_video_metadata(self, file_content: bytes, file_name: str) -> str:
        """
        Analisa metadados de vídeo
        
        Args:
            file_content: Conteúdo do vídeo em bytes
            file_name: Nome do arquivo
            
        Returns:
            Informações sobre o vídeo
        """
        try:
            import struct
            
            # Informações básicas
            size_mb = len(file_content) / (1024 * 1024)
            ext = file_name.lower().split('.')[-1]
            
            info = f"[VÍDEO]\n"
            info += f"Nome: {file_name}\n"
            info += f"Formato: {ext.upper()}\n"
            info += f"Tamanho: {size_mb:.2f} MB\n"
            
            # Tenta detectar resolução e duração (simplificado)
            if ext == 'mp4':
                info += "Tipo: Vídeo MP4\n"
            elif ext == 'avi':
                info += "Tipo: Vídeo AVI\n"
            elif ext == 'mov':
                info += "Tipo: Vídeo QuickTime\n"
            else:
                info += f"Tipo: Vídeo {ext.upper()}\n"
            
            info += "\nNota: Este é um arquivo de vídeo. Para análise completa do conteúdo, seria necessário extrair frames e processá-los individualmente."
            
            return info
            
        except Exception as e:
            return f"[VÍDEO] {file_name} - Tamanho: {len(file_content) / (1024 * 1024):.2f} MB"
    
    def analyze_executable(self, file_content: bytes, file_name: str) -> str:
        """
        Analisa arquivo executável ou binário
        
        Args:
            file_content: Conteúdo do arquivo em bytes
            file_name: Nome do arquivo
            
        Returns:
            Informações sobre o executável
        """
        try:
            size_mb = len(file_content) / (1024 * 1024)
            ext = file_name.lower().split('.')[-1]
            
            info = f"[EXECUTÁVEL/BINÁRIO]\n"
            info += f"Nome: {file_name}\n"
            info += f"Extensão: {ext.upper()}\n"
            info += f"Tamanho: {size_mb:.2f} MB\n"
            
            # Detecta tipo de executável
            if file_content.startswith(b'MZ'):
                info += "Tipo: Executável Windows (PE)\n"
            elif file_content.startswith(b'\x7fELF'):
                info += "Tipo: Executável Linux (ELF)\n"
            elif file_content.startswith(b'\xca\xfe\xba\xbe'):
                info += "Tipo: Executável macOS (Mach-O)\n"
            else:
                info += f"Tipo: Arquivo binário {ext.upper()}\n"
            
            # Tenta extrair strings legíveis
            try:
                readable_strings = []
                current_string = []
                for byte in file_content[:10000]:  # Primeiros 10KB
                    if 32 <= byte <= 126:  # Caracteres ASCII imprimíveis
                        current_string.append(chr(byte))
                    else:
                        if len(current_string) >= 4:
                            readable_strings.append(''.join(current_string))
                        current_string = []
                
                if readable_strings:
                    info += f"\nStrings encontradas (primeiras 10): {', '.join(readable_strings[:10])}"
            except:
                pass
            
            return info
            
        except Exception as e:
            return f"[BINÁRIO] {file_name} - Tamanho: {len(file_content) / (1024 * 1024):.2f} MB"
    
    def analyze_archive(self, file_content: bytes, file_name: str) -> str:
        """
        Analisa arquivo compactado
        
        Args:
            file_content: Conteúdo do arquivo em bytes
            file_name: Nome do arquivo
            
        Returns:
            Informações sobre o arquivo
        """
        try:
            import zipfile
            
            size_mb = len(file_content) / (1024 * 1024)
            ext = file_name.lower().split('.')[-1]
            
            info = f"[ARQUIVO COMPACTADO]\n"
            info += f"Nome: {file_name}\n"
            info += f"Formato: {ext.upper()}\n"
            info += f"Tamanho: {size_mb:.2f} MB\n"
            
            # Tenta listar conteúdo se for ZIP
            if ext == 'zip':
                try:
                    with zipfile.ZipFile(io.BytesIO(file_content)) as zf:
                        files = zf.namelist()
                        info += f"\nArquivos contidos ({len(files)} total):\n"
                        for f in files[:20]:  # Primeiros 20
                            info += f"  - {f}\n"
                        if len(files) > 20:
                            info += f"  ... e mais {len(files) - 20} arquivos\n"
                except:
                    info += "\nNão foi possível listar o conteúdo do arquivo ZIP.\n"
            
            return info
            
        except Exception as e:
            return f"[ARQUIVO] {file_name} - Tamanho: {len(file_content) / (1024 * 1024):.2f} MB"

    def extract_text_from_file(self, file_content: bytes, file_type: str, file_name: str = "") -> str:
        """
        Extrai texto/informação de QUALQUER tipo de arquivo
        
        Args:
            file_content: Conteúdo do arquivo em bytes
            file_type: Tipo do arquivo (pdf, word, excel, powerpoint, text, image, video, etc.)
            file_name: Nome do arquivo (opcional, usado para análise)
        
        Returns:
            Texto extraído ou análise do arquivo
        """
        # Documentos de texto tradicionais
        if file_type == 'pdf':
            return self.extract_text_from_pdf(file_content)
        elif file_type == 'word':
            return self.extract_text_from_docx(file_content)
        elif file_type == 'excel':
            return self.extract_text_from_xlsx(file_content)
        elif file_type == 'powerpoint':
            return self.extract_text_from_pptx(file_content)
        elif file_type == 'text':
            try:
                return file_content.decode('utf-8', errors='ignore')
            except:
                return file_content.decode('latin-1', errors='ignore')
        
        # Imagens - usa Claude Vision
        elif file_type == 'image':
            return self.analyze_image(file_content, file_name)
        
        # Vídeos - analisa metadados
        elif file_type == 'video':
            return self.analyze_video_metadata(file_content, file_name)
        
        # Áudio - metadados
        elif file_type == 'audio':
            size_mb = len(file_content) / (1024 * 1024)
            return f"[ÁUDIO]\nNome: {file_name}\nTamanho: {size_mb:.2f} MB\nTipo: Arquivo de áudio\n\nNota: Este é um arquivo de áudio. Para análise completa, seria necessário transcrição de fala."
        
        # Arquivos compactados
        elif file_type == 'archive':
            return self.analyze_archive(file_content, file_name)
        
        # Executáveis e binários
        elif file_type == 'executable':
            return self.analyze_executable(file_content, file_name)
        
        # CAD, Design, etc.
        elif file_type in ['cad', 'design']:
            size_mb = len(file_content) / (1024 * 1024)
            return f"[{file_type.upper()}]\nNome: {file_name}\nTamanho: {size_mb:.2f} MB\nTipo: Arquivo de design/CAD\n\nNota: Este é um arquivo de design profissional. Contém dados técnicos e gráficos."
        
        # Qualquer outro tipo de arquivo
        else:
            size_mb = len(file_content) / (1024 * 1024)
            ext = file_name.split('.')[-1] if '.' in file_name else file_type
            
            # Tenta detectar se é texto
            try:
                text_content = file_content.decode('utf-8', errors='strict')
                if len(text_content) > 0 and len(text_content) < 1000000:  # Menos de 1MB de texto
                    return f"[ARQUIVO DE TEXTO - {ext.upper()}]\n\n{text_content}"
            except:
                pass
            
            # Se não for texto, retorna informações básicas
            return f"[ARQUIVO GENÉRICO]\nNome: {file_name}\nExtensão: {ext.upper()}\nTamanho: {size_mb:.2f} MB\nTipo: {file_type}\n\nNota: Este arquivo foi detectado mas não possui um analisador específico. Informações básicas foram extraídas."

    def summarize_document(
        self,
        file_name: str,
        file_content: bytes,
        file_type: str,
        detail_level: str = 'medium'
    ) -> Dict:
        """
        Resume um documento usando Claude 3

        Args:
            file_name: Nome do arquivo
            file_content: Conteúdo do arquivo em bytes
            file_type: Tipo do arquivo (pdf, word, excel, etc)
            detail_level: Nível de detalhe (low, medium, high)

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

            # Define prompt baseado no nível de detalhe
            if detail_level == 'low':
                print(f"🎯 Claude: Usando prompt de nível BAIXO (conciso)")
                system_prompt = """Você é um assistente focado em brevidade.
Sua tarefa é criar resumos extremamente concisos e diretos.

Diretrizes:
- Identifique APENAS os 3 pontos mais críticos
- Use no máximo 3-4 frases
- Ignore detalhes secundários
- Seja direto ao ponto
"""
            elif detail_level == 'high':
                print(f"🎯 Claude: Usando prompt de nível ALTO (detalhado)")
                system_prompt = """Você é um analista detalhista.
Sua tarefa é criar resumos abrangentes e profundos.

Diretrizes:
- Cubra todos os aspectos importantes do documento
- Inclua detalhes técnicos, datas específicas e valores
- Explique o contexto e as nuances
- Use formatação estruturada com seções se necessário
- Não omita informações relevantes
"""
            else: # medium (padrão)
                print(f"🎯 Claude: Usando prompt de nível MÉDIO (equilibrado)")
                system_prompt = """Você é um assistente especializado em analisar e resumir documentos corporativos.
Sua tarefa é criar resumos concisos, informativos e bem estruturados.

Diretrizes:
- Identifique os pontos principais e informações mais relevantes
- Use formatação markdown para melhor legibilidade
- Destaque dados importantes, datas, valores e decisões
- Seja objetivo e direto
- Use bullets quando apropriado
"""
            
            # User prompt (common across detail levels)
            user_prompt = f"""Analise o seguinte documento e forneça:

1. **Resumo Executivo**: Um resumo conciso do conteúdo principal (2-3 parágrafos)
2. **Pontos Principais**: Liste os 5-7 pontos mais importantes do documento
3. **Informações Chave**: Dados, números ou fatos relevantes mencionados
4. **Categoria**: Qual o tipo/categoria deste documento (ex: relatório, apresentação, planilha, etc.)

Documento: {file_name}

Por favor, forneça a análise em português do Brasil, de forma clara e profissional."""

            messages = []
            
            # Lógica para PDF direto (se extração de texto falhou ou retornou None)
            if (extracted_text is None or len(extracted_text) < 50) and file_type == 'pdf':
                print("🔄 Usando fallback de PDF direto para Claude (Vision/PDF)...")
                pdf_base64 = base64.b64encode(file_content).decode('utf-8')
                
                messages = [
                    {
                        "role": "user",
                        "content": [
                            {
                                "type": "document",
                                "source": {
                                    "type": "base64",
                                    "media_type": "application/pdf",
                                    "data": pdf_base64
                                }
                            },
                            {
                                "type": "text",
                                "text": base_prompt
                            }
                        ]
                    }
                ]
                # Define um texto placeholder para contagem de caracteres
                extracted_text = "[PDF processado diretamente via Claude Vision]"
                
            else:
                # Fluxo normal de texto
                if extracted_text is None:
                    extracted_text = ""
                    
                # Limita o texto se for muito grande
                max_chars = 100000
                if len(extracted_text) > max_chars:
                    extracted_text = extracted_text[:max_chars] + "\n\n[... documento truncado ...]"
                
                full_prompt = base_prompt + f"\n\nConteúdo:\n{extracted_text}"
                
                messages = [
                    {"role": "user", "content": full_prompt}
                ]

            # Chama Claude API
            message = self.client.messages.create(
                model=self.model,
                max_tokens=2000,
                messages=messages
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
