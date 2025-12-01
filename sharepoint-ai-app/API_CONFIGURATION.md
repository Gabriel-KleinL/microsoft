# Documentação: Configuração das APIs de IA

## Visão Geral

Esta aplicação SharePoint + IA suporta dois provedores de Inteligência Artificial:

1. **Claude AI (Anthropic)** - Modelo padrão: `claude-sonnet-4-5-20250929`
2. **OpenAI (ChatGPT)** - Modelo padrão: `gpt-4`

O sistema foi projetado com uma **arquitetura modular** que permite alternar entre os provedores apenas modificando uma variável de ambiente.

---

## Arquitetura do Sistema

```
┌─────────────────────────────────────────────────┐
│           Frontend (JavaScript)                  │
│  - Interface do usuário                          │
│  - Seleção de documentos                         │
│  - Chat interativo                               │
└────────────────┬────────────────────────────────┘
                 │
                 │ HTTP/REST
                 ▼
┌─────────────────────────────────────────────────┐
│         Backend Flask (Python)                   │
│  ┌───────────────────────────────────────────┐  │
│  │   app.py - Rotas e Lógica Principal       │  │
│  └───────────────────────────────────────────┘  │
│                     │                            │
│    ┌────────────────┼────────────────┐           │
│    ▼                ▼                ▼           │
│  ┌────────┐   ┌──────────┐   ┌──────────┐       │
│  │ Config │   │ MS Graph │   │ AI Service│       │
│  └────────┘   └──────────┘   └──────────┘       │
│                                    │             │
│                       ┌────────────┴─────────┐   │
│                       ▼                      ▼   │
│              ┌─────────────────┐   ┌─────────────────┐
│              │ ClaudeAIService │   │ OpenAIService   │
│              └─────────────────┘   └─────────────────┘
└─────────────────────────────────────────────────┘
                       │                      │
                       ▼                      ▼
              ┌─────────────────┐   ┌─────────────────┐
              │   Anthropic     │   │     OpenAI      │
              │   Claude API    │   │   ChatGPT API   │
              └─────────────────┘   └─────────────────┘
```

---

## 1. Arquivo de Configuração (.env)

### Localização
```
sharepoint-ai-app/backend/.env
```

### Estrutura do Arquivo

```env
# ============================================
# MICROSOFT AZURE AD / GRAPH API
# ============================================
MICROSOFT_CLIENT_ID=seu-client-id-aqui
MICROSOFT_CLIENT_SECRET=seu-client-secret-aqui
MICROSOFT_TENANT_ID=seu-tenant-id-aqui
MICROSOFT_REDIRECT_URI=http://localhost:5000/auth/callback

# ============================================
# SHAREPOINT
# ============================================
SHAREPOINT_SITE_URL=seu-site-sharepoint.com

# ============================================
# CLAUDE AI (ANTHROPIC)
# ============================================
CLAUDE_API_KEY=sk-ant-api03-xxxxx
CLAUDE_MODEL=claude-sonnet-4-5-20250929

# ============================================
# OPENAI (CHATGPT)
# ============================================
OPENAI_API_KEY=sk-xxxxx
OPENAI_MODEL=gpt-4

# ============================================
# PROVEDOR DE IA ATIVO
# ============================================
# Valores possíveis: 'claude' ou 'openai'
AI_PROVIDER=claude

# ============================================
# FLASK
# ============================================
SECRET_KEY=sua-chave-secreta-aqui
FLASK_ENV=development
```

---

## 2. Classe de Configuração (config.py)

### Localização
```
sharepoint-ai-app/backend/config.py
```

### Funcionamento

```python
class Config:
    # Microsoft Azure AD / Graph API
    MICROSOFT_CLIENT_ID = os.getenv('MICROSOFT_CLIENT_ID')
    MICROSOFT_CLIENT_SECRET = os.getenv('MICROSOFT_CLIENT_SECRET')
    MICROSOFT_TENANT_ID = os.getenv('MICROSOFT_TENANT_ID')

    # SharePoint
    SHAREPOINT_SITE_URL = os.getenv('SHAREPOINT_SITE_URL')

    # Claude API
    CLAUDE_API_KEY = os.getenv('CLAUDE_API_KEY')
    CLAUDE_MODEL = os.getenv('CLAUDE_MODEL', 'claude-sonnet-4-5-20250929')

    # OpenAI API
    OPENAI_API_KEY = os.getenv('OPENAI_API_KEY')
    OPENAI_MODEL = os.getenv('OPENAI_MODEL', 'gpt-4')

    # Provedor de IA (claude ou openai)
    AI_PROVIDER = os.getenv('AI_PROVIDER', 'claude').lower()
```

**Validação Automática**: O método `Config.validate()` verifica se:
- Todas as configurações obrigatórias estão presentes
- O `AI_PROVIDER` é válido ('claude' ou 'openai')
- A API key do provedor selecionado está configurada

---

## 3. Inicialização dos Serviços de IA

### Localização
```
sharepoint-ai-app/backend/app.py (linhas 49-67)
```

### Lógica de Seleção

```python
# Inicializa serviços de IA baseado no provedor configurado
ai_service = None
ai_provider_name = Config.AI_PROVIDER

if ai_provider_name == 'claude':
    ai_service = ClaudeAIService(
        api_key=Config.CLAUDE_API_KEY,
        model=Config.CLAUDE_MODEL
    )
    print(f"✓ Usando Claude AI: {Config.CLAUDE_MODEL}")

elif ai_provider_name == 'openai':
    ai_service = OpenAIService(
        api_key=Config.OPENAI_API_KEY,
        model=Config.OPENAI_MODEL
    )
    print(f"✓ Usando OpenAI: {Config.OPENAI_MODEL}")

else:
    print(f"✗ Provedor de IA inválido: {ai_provider_name}")
```

O serviço ativo é acessado globalmente através da função:
```python
def get_ai_service():
    """Retorna o serviço de IA configurado"""
    return ai_service
```

---

## 4. Claude AI Service

### Localização
```
sharepoint-ai-app/backend/services/claude_ai.py
```

### Funcionalidades

#### Inicialização
```python
class ClaudeAIService:
    def __init__(self, api_key: str, model: str = "claude-sonnet-4-5-20250929"):
        self.client = anthropic.Anthropic(api_key=api_key)
        self.model = model
```

#### Métodos Principais

1. **`summarize_document(file_content, file_type, file_name)`**
   - Resume documentos (PDF, DOCX, XLSX, CSV, PPTX, TXT)
   - Usa extração de texto específica para cada tipo
   - Retorna Dict com resumo estruturado

2. **`extract_text_from_file(file_content, file_type, file_name)`**
   - Extrai texto de arquivos binários
   - Suporte a múltiplos formatos
   - Métodos cascata para PDFs (PyPDF2 → pdfplumber)

3. **`extract_text_from_pdf(file_content)`**
   - **Método 1**: PyPDF2 (rápido)
   - **Método 2**: pdfplumber (robusto)
   - **Fallback**: Mensagem informativa

4. **`extract_text_from_xlsx(file_content)`**
   - Detecta automaticamente CSV vs XLSX
   - Suporte a delimitadores: `,`, `;`, `\t`
   - Processa múltiplas planilhas

5. **`generate_combined_summary(individual_summaries, file_names)`**
   - Consolida múltiplos resumos
   - Identifica temas comuns
   - Organiza por tópicos

#### Exemplo de Uso na Aplicação

```python
# Em app.py:444-543 (/api/ai/summarize)
summary = get_ai_service().summarize_document(
    file_content=file_content,
    file_type=doc_info['type'],
    file_name=doc_info['name']
)

# Retorna:
{
    'success': True,
    'file_name': 'documento.pdf',
    'file_type': 'pdf',
    'summary': 'Resumo do documento...',
    'char_count': 5000,
    'model_used': 'claude-sonnet-4-5-20250929'
}
```

---

## 5. OpenAI Service

### Localização
```
sharepoint-ai-app/backend/services/openai_service.py
```

### Funcionalidades

#### Inicialização
```python
class OpenAIService:
    def __init__(self, api_key: str, model: str = "gpt-4"):
        self.client = OpenAI(api_key=api_key)
        self.model = model
        # Reutiliza extração de texto do ClaudeAI
        self._text_extractor = ClaudeAIService(api_key="dummy", model="dummy")
```

#### Métodos Principais

1. **`chat(messages, max_tokens, temperature)`**
   - Interface direta com ChatGPT API
   - Suporta conversação multi-turn
   - Configurável (tokens, temperatura)

2. **`summarize_document(file_name, file_content, file_type)`**
   - Resume documentos usando ChatGPT
   - **Limite de conteúdo**: 3000 caracteres (evita context_length_exceeded)
   - **Max tokens resposta**: 800 tokens
   - Retorna Dict (compatível com ClaudeAI)

3. **`consolidate_summaries(summaries, document_count)`**
   - Consolida múltiplos resumos
   - Cria resumo executivo integrado

4. **`answer_question(question, documents_context, conversation_history)`**
   - Responde perguntas baseadas em documentos
   - RAG (Retrieval Augmented Generation)
   - Mantém contexto da conversa

5. **`extract_text_from_file(file_content, file_type, file_name)`**
   - **Reutiliza a lógica do ClaudeAIService**
   - Extração de texto é independente da IA
   - Evita duplicação de código

#### Exemplo de Uso na Aplicação

```python
# Em app.py:444-543 (/api/ai/summarize)
summary = get_ai_service().summarize_document(
    file_content=file_content,
    file_type=doc_info['type'],
    file_name=doc_info['name']
)

# Retorna (mesmo formato do Claude):
{
    'success': True,
    'file_name': 'documento.pdf',
    'file_type': 'pdf',
    'summary': 'Resumo do documento...',
    'char_count': 5000,
    'model_used': 'gpt-4'
}
```

---

## 6. Como Alternar Entre Claude e OpenAI

### Método 1: Variável de Ambiente (Recomendado)

Edite o arquivo `.env`:

```env
# Para usar Claude AI
AI_PROVIDER=claude

# Para usar OpenAI
AI_PROVIDER=openai
```

Reinicie o servidor Flask:
```bash
pkill -f "python.*backend/app.py"
./venv/bin/python backend/app.py
```

### Método 2: Variável de Ambiente no Terminal

```bash
# Claude AI
export AI_PROVIDER=claude
./venv/bin/python backend/app.py

# OpenAI
export AI_PROVIDER=openai
./venv/bin/python backend/app.py
```

---

## 7. Endpoints da API que Usam IA

### 1. `/api/ai/summarize` (POST)

**Função**: Resume um ou múltiplos documentos

**Request**:
```json
{
  "document_ids": [
    {
      "id": "doc-id-1",
      "driveId": "drive-id-1",
      "name": "documento.pdf",
      "type": "pdf"
    }
  ]
}
```

**Response**:
```json
{
  "success": true,
  "summaries": [
    {
      "success": true,
      "file_name": "documento.pdf",
      "file_type": "pdf",
      "summary": "## Resumo do Documento\n\n**Assunto Principal**: ...",
      "char_count": 5000,
      "model_used": "claude-sonnet-4-5-20250929"
    }
  ],
  "consolidated_summary": "# Resumo Consolidado\n\n...",
  "total_processed": 1,
  "cache_hits": 0
}
```

**Provedor Usado**: Definido por `AI_PROVIDER` (Claude ou OpenAI)

---

### 2. `/api/ai/chat` (POST)

**Função**: Chat com IA usando RAG (Retrieval Augmented Generation)

**Request**:
```json
{
  "message": "Qual o orçamento do projeto X?",
  "conversation_history": [
    {"role": "user", "content": "Olá"},
    {"role": "assistant", "content": "Olá! Como posso ajudar?"}
  ],
  "project_instructions": "Você é um assistente especializado em finanças."
}
```

**Response**:
```json
{
  "success": true,
  "response": "Baseado nos documentos analisados, o orçamento do projeto X é...",
  "sources": [
    {
      "name": "orcamento_2024.xlsx",
      "type": "excel",
      "webUrl": "https://sharepoint.com/..."
    }
  ],
  "documents_analyzed": 2,
  "documents_found": 5
}
```

**Funcionamento**:
1. Busca documentos relevantes no SharePoint
2. Extrai conteúdo dos 3 primeiros documentos (máx 10.000 caracteres cada)
3. Constrói contexto para IA
4. Envia pergunta + contexto para IA (Claude ou OpenAI)
5. Retorna resposta com fontes citadas

**Diferença entre Provedores**:
- **Claude**: Usa `client.messages.create()` com system prompt separado
- **OpenAI**: Usa `answer_question()` que injeta context no system prompt

---

## 8. Gestão de Limites de Tokens

### Claude AI
- **Limite do modelo**: ~200.000 tokens (Claude Sonnet 4.5)
- **Configuração atual**: Sem limite rígido (modelo é generoso)
- **Context window**: Muito amplo, raramente atinge limite

### OpenAI (GPT-4)
- **Limite do modelo**: 8.192 tokens
- **Problema anterior**: `context_length_exceeded` errors
- **Solução implementada**:

```python
# Em openai_service.py:112-138
max_content_chars = 3000  # ~1500 tokens

# Distribui tokens:
# - System prompt: ~200 tokens
# - User prompt structure: ~150 tokens
# - Document content: ~1500 tokens (3000 chars)
# - Response: 800 tokens
# Total: ~2650 tokens (bem abaixo de 8192)

if len(content_str) > max_content_chars:
    content_preview = content_str[:max_content_chars] + "\n\n[... documento truncado ...]"
```

---

## 9. Extração de Texto de Documentos

### Tipos de Arquivo Suportados

| Tipo | Extensão | Biblioteca | Método |
|------|----------|------------|--------|
| PDF | .pdf | PyPDF2, pdfplumber | `extract_text_from_pdf()` |
| Word | .docx | python-docx | `extract_text_from_word()` |
| Excel | .xlsx | openpyxl | `extract_text_from_xlsx()` |
| CSV | .csv | csv (built-in) | `extract_text_from_xlsx()` |
| PowerPoint | .pptx | python-pptx | `extract_text_from_ppt()` |
| Texto | .txt | decode utf-8 | Direct decode |

### Estratégia de Extração Cascata (PDF)

```python
# Método 1: PyPDF2 (rápido)
try:
    pdf_reader = PyPDF2.PdfReader(file_content)
    text = extract_text_from_pages()
    if len(text) > 50:
        return text  # Sucesso!
except:
    pass

# Método 2: pdfplumber (robusto)
try:
    with pdfplumber.open(file_content) as pdf:
        text = extract_text_from_pages()
        if len(text) > 50:
            return text  # Sucesso!
except:
    pass

# Fallback: Mensagem informativa
return "[PDF - Extração Limitada] ..."
```

### Detecção Automática CSV vs XLSX

```python
# Tenta primeiro como CSV (arquivo de texto)
try:
    csv_text = file_content.decode('utf-8')
    if ',' in csv_text or ';' in csv_text or '\t' in csv_text:
        # Detecta delimitador
        delimiter = detect_delimiter(csv_text)
        # Processa como CSV
        return process_csv(csv_text, delimiter)
except:
    pass

# Se falhar, tenta como XLSX (arquivo binário)
workbook = openpyxl.load_workbook(file_content)
return process_xlsx(workbook)
```

---

## 10. Cache de Resumos

### Sistema de Cache

**Tipo**: SimpleCache (em memória)
**Timeout**: 30 minutos (1800 segundos)

### Chaves de Cache

```python
# Resumo individual
cache_key = f"summary_{doc_id}_{drive_id}"

# Resumo consolidado
cache_key = f"consolidated_{sorted_doc_ids_joined}"

# Documentos do SharePoint
cache_key = f"docs_{user_email}_{folder_path}"

# Busca
cache_key = f"search_{user_email}_{query}_{max_results}"
```

### Benefícios

1. **Performance**: Evita re-processar documentos já resumidos
2. **Custo**: Reduz chamadas às APIs pagas (Claude/OpenAI)
3. **Experiência**: Respostas instantâneas para documentos em cache

### Limpeza Manual

```bash
# Via API
POST /api/cache/clear
```

---

## 11. Interface Unificada

### Ambos os serviços implementam a mesma interface:

```python
class IAService:
    def summarize_document(file_content, file_type, file_name) -> Dict
    def extract_text_from_file(file_content, file_type, file_name) -> str
    def generate_combined_summary(summaries, file_names) -> str
```

### Vantagens da Arquitetura

1. **Intercambiabilidade**: Trocar provedor sem alterar código
2. **Testabilidade**: Fácil criar mocks para testes
3. **Manutenibilidade**: Mudanças isoladas em cada serviço
4. **Escalabilidade**: Fácil adicionar novos provedores (ex: Google Gemini)

---

## 12. Fluxo Completo de Resumo

```
┌─────────────────────────────────────────────────────┐
│ 1. Usuário seleciona documento(s) no frontend       │
└────────────────┬────────────────────────────────────┘
                 │
                 ▼
┌─────────────────────────────────────────────────────┐
│ 2. Frontend envia POST /api/ai/summarize            │
│    Body: { document_ids: [...] }                    │
└────────────────┬────────────────────────────────────┘
                 │
                 ▼
┌─────────────────────────────────────────────────────┐
│ 3. Backend verifica cache                           │
│    - Cache hit? Retorna resumo salvo                │
│    - Cache miss? Continua...                        │
└────────────────┬────────────────────────────────────┘
                 │
                 ▼
┌─────────────────────────────────────────────────────┐
│ 4. Backend baixa arquivo do SharePoint              │
│    graph_service.download_file_content()            │
└────────────────┬────────────────────────────────────┘
                 │
                 ▼
┌─────────────────────────────────────────────────────┐
│ 5. Extrai texto do arquivo                          │
│    ai_service.extract_text_from_file()              │
│    - PDF: PyPDF2 → pdfplumber                       │
│    - DOCX: python-docx                              │
│    - XLSX/CSV: openpyxl / csv                       │
└────────────────┬────────────────────────────────────┘
                 │
                 ▼
┌─────────────────────────────────────────────────────┐
│ 6. Envia para IA (Claude ou OpenAI)                 │
│    ai_service.summarize_document()                  │
│                                                      │
│    Claude:                                           │
│    - anthropic.messages.create()                    │
│    - Max tokens: 2000                               │
│                                                      │
│    OpenAI:                                           │
│    - openai.chat.completions.create()               │
│    - Max tokens: 800                                │
│    - Content limit: 3000 chars                      │
└────────────────┬────────────────────────────────────┘
                 │
                 ▼
┌─────────────────────────────────────────────────────┐
│ 7. Salva resumo no cache (30 min)                   │
└────────────────┬────────────────────────────────────┘
                 │
                 ▼
┌─────────────────────────────────────────────────────┐
│ 8. Se múltiplos docs, gera resumo consolidado       │
│    ai_service.generate_combined_summary()           │
└────────────────┬────────────────────────────────────┘
                 │
                 ▼
┌─────────────────────────────────────────────────────┐
│ 9. Retorna JSON com resumos                         │
│    {                                                 │
│      summaries: [...],                              │
│      consolidated_summary: "...",                   │
│      cache_hits: 0                                  │
│    }                                                 │
└─────────────────────────────────────────────────────┘
```

---

## 13. Logs e Debugging

### Logs de Inicialização

```bash
✓ Configurações validadas com sucesso
✓ Claude AI Service inicializado - Modelo: claude-sonnet-4-5-20250929
✓ Usando Claude AI: claude-sonnet-4-5-20250929

# Ou

✓ OpenAI Service inicializado - Modelo: gpt-4
✓ Usando OpenAI: gpt-4
```

### Logs de Extração de PDF

```bash
📄 PyPDF2: Extraído 5420 caracteres de 12 páginas
✅ PyPDF2 sucesso!

# Se PyPDF2 falhar:
❌ PyPDF2 falhou: [erro]
📄 pdfplumber: Extraído 5380 caracteres de 12 páginas
✅ pdfplumber sucesso!
```

### Logs de Cache

```bash
✓ Cache hit para resumo: documento.pdf
✓ Cache salvo para resumo: documento.pdf
✓ Cache hit para resumo consolidado
```

### Logs de Chat (RAG)

```bash
==================================================================
║ 🤖 SOPH-IA - PROCESSAMENTO DE CHAT
==================================================================
👤 Usuário: user@example.com
📝 Pergunta: Qual o orçamento do projeto?
⏰ Timestamp: 2025-01-15 14:30:00
💬 Histórico: 2 mensagens anteriores
==================================================================

──────────────────────────────────────────────────────────────────
🔍 ETAPA 1: BUSCANDO DOCUMENTOS NO SHAREPOINT
──────────────────────────────────────────────────────────────────
📊 Query de busca: 'Qual o orçamento do projeto?'
📈 Máximo de resultados: 10

✅ Busca concluída: 5 documentos encontrados
📄 Documentos encontrados:
   1. orcamento_2024.xlsx (excel) - Drive: Documentos Compartilhados
   2. projeto_x_plano.docx (word) - Drive: Documentos Compartilhados

──────────────────────────────────────────────────────────────────
📄 ETAPA 2: EXTRAINDO CONTEÚDO DOS DOCUMENTOS
──────────────────────────────────────────────────────────────────
📊 Documentos a analisar: 3

   [1/3] 📄 orcamento_2024.xlsx
       Tipo: excel
       Tamanho: 45832 bytes
       ⬇️  Baixando arquivo...
       ✓ Download concluído (45832 bytes)
       📝 Extraindo texto...
       ✅ Texto extraído: 3,240 caracteres
       📊 Preview: === Planilha: Orçamento 2024 ===...

──────────────────────────────────────────────────────────────────
🧠 ETAPA 3: CONSTRUINDO CONTEXTO PARA CLAUDE AI
──────────────────────────────────────────────────────────────────
📊 Contexto construído: 8,420 caracteres
📄 Documentos no contexto: 3
   1. orcamento_2024.xlsx - 3,240 chars
   2. projeto_x_plano.docx - 5,180 chars

──────────────────────────────────────────────────────────────────
✨ ETAPA 4: GERANDO RESPOSTA COM CLAUDE AI
──────────────────────────────────────────────────────────────────
📤 Parâmetros da requisição:
   Modelo: claude-sonnet-4-5-20250929
   Max tokens: 2000
   System prompt: 1,280 caracteres
   Mensagens no histórico: 3
   Pergunta: "Qual o orçamento do projeto?"

⏳ Aguardando resposta de CLAUDE AI...
✅ Resposta recebida!
   ⏱️  Tempo de resposta: 2.34s
   📝 Tamanho da resposta: 543 caracteres
   📊 Tokens usados: ~450 input / ~180 output
   💬 Preview: Baseado no documento **orcamento_2024.xlsx**...

==================================================================
✅ PROCESSAMENTO COMPLETO - SOPH-IA RESPONDEU!
==================================================================
```

---

## 14. Dependências Necessárias

### requirements.txt

```txt
# Framework Web
Flask==3.0.0
flask-cors==4.0.0
Flask-Caching==2.1.0

# Microsoft Graph / Azure AD
msal==1.26.0
requests==2.31.0

# Claude AI (Anthropic)
anthropic>=0.40.0

# OpenAI
openai>=1.0.0

# Processamento de Documentos
PyPDF2==3.0.1
pdfplumber==0.11.8
python-docx==1.1.0
python-pptx==0.6.23
openpyxl==3.1.2

# Variáveis de Ambiente
python-dotenv==1.0.0

# Utilitários
certifi==2023.11.17
```

### Instalação

```bash
cd sharepoint-ai-app
python3 -m venv venv
source venv/bin/activate  # Linux/Mac
# ou
venv\Scripts\activate  # Windows

pip install -r backend/requirements.txt
```

---

## 15. Troubleshooting

### Erro: "AI_PROVIDER inválido"

**Causa**: Valor incorreto na variável `AI_PROVIDER`

**Solução**:
```env
# Valores válidos: 'claude' ou 'openai' (minúsculas)
AI_PROVIDER=claude
```

### Erro: "Configurações ausentes: CLAUDE_API_KEY"

**Causa**: API key não configurada no .env

**Solução**:
```env
CLAUDE_API_KEY=sk-ant-api03-seu-token-aqui
```

### Erro: "context_length_exceeded" (OpenAI)

**Causa**: Documento muito grande

**Solução**: Já implementada no código (limite de 3000 chars). Se persistir:
```python
# Em openai_service.py, reduza ainda mais:
max_content_chars = 2000  # em vez de 3000
```

### Erro: PDF não extrai texto

**Causa**: PDF contém imagens ou está protegido

**Solução**: O sistema já tenta 2 métodos (PyPDF2 + pdfplumber). Se ambos falharem:
- Converta o PDF para formato pesquisável (OCR)
- Use ferramentas como Adobe Acrobat para remover proteção

### Cache não limpa

**Solução**:
```bash
# Via API
curl -X POST http://localhost:5000/api/cache/clear

# Ou reinicie o servidor
pkill -f "python.*backend/app.py"
./venv/bin/python backend/app.py
```

---

## 16. Segurança

### API Keys

- **NUNCA** commite o arquivo `.env` no Git
- Use `.env.example` como template
- Rotacione as API keys periodicamente
- Use variáveis de ambiente em produção

### Token Management

- Tokens são armazenados em sessão server-side (Flask-Session)
- Auto-renovação de tokens expirados
- Logout limpa todos os tokens

### Permissões do SharePoint

```python
SCOPES = [
    'User.Read',        # Ler perfil do usuário
    'Files.Read.All',   # Ler arquivos
    'Sites.Read.All'    # Ler sites SharePoint
]
```

---

## 17. Performance

### Otimizações Implementadas

1. **Cache de Resumos**: 30 minutos (reduz chamadas à IA)
2. **Cache de Documentos**: 5 minutos (reduz chamadas ao SharePoint)
3. **Cache de Site ID**: 30 minutos (evita lookups repetidos)
4. **Limite de Documentos no Chat**: Máximo 3 documentos analisados
5. **Truncamento de Conteúdo**: Máximo 10.000 caracteres por documento no chat

### Métricas Típicas

- **Tempo de resumo (1 doc)**: 2-5 segundos
- **Tempo de chat RAG**: 3-8 segundos
- **Cache hit**: < 100ms
- **Download de arquivo**: 200ms - 2s (depende do tamanho)

---

## Conclusão

Este sistema oferece uma integração flexível e robusta entre SharePoint e múltiplos provedores de IA. A arquitetura modular permite:

- Fácil troca entre Claude AI e OpenAI
- Adição de novos provedores no futuro
- Manutenção independente de cada serviço
- Otimização de custos através de cache
- Suporte robusto a múltiplos formatos de arquivo

Para suporte adicional, consulte os comentários no código ou entre em contato com a equipe de desenvolvimento.
