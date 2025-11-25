# 📄 SharePoint + Claude AI Integration

Aplicação web que integra SharePoint com Claude AI (Anthropic) para resumir documentos automaticamente.

## 🚀 Funcionalidades

- ✅ **Autenticação Microsoft OAuth 2.0** - Login seguro com conta Microsoft
- ✅ **Integração com SharePoint** - Acesso a documentos via Microsoft Graph API
- ✅ **IA Generativa** - Resumos inteligentes usando Claude API (Anthropic)
- ✅ **Múltiplos Formatos** - Suporte para PDF, Word (.docx), Excel (.xlsx), PowerPoint (.pptx)
- ✅ **Interface Moderna** - Design profissional e responsivo
- ✅ **Resumo Consolidado** - Análise conjunta de múltiplos documentos

## 📋 Pré-requisitos

- Python 3.8 ou superior
- Conta Microsoft 365 com acesso ao SharePoint
- Conta Anthropic (para Claude API)
- Acesso ao Azure Portal para registro de aplicativo

## 🛠️ Instalação

### 1. Clone o repositório

```bash
git clone <seu-repositorio>
cd sharepoint-ai-app
```

### 2. Crie um ambiente virtual Python

```bash
# Linux/Mac
python3 -m venv venv
source venv/bin/activate

# Windows
python -m venv venv
venv\Scripts\activate
```

### 3. Instale as dependências

```bash
cd backend
pip install -r requirements.txt
```

### 4. Configure as variáveis de ambiente

```bash
# Copie o arquivo de exemplo
cp ../.env.example ../.env

# Edite o arquivo .env com suas credenciais
nano ../.env  # ou use seu editor preferido
```

## 🔑 Configuração do Azure AD (Microsoft)

### Passo 1: Registrar Aplicativo no Azure Portal

1. Acesse o [Azure Portal](https://portal.azure.com/)
2. Vá para **Azure Active Directory** → **App registrations** → **New registration**

### Passo 2: Configurar o Aplicativo

**Nome do aplicativo:**
```
SharePoint AI Integration
```

**Tipos de conta suportados:**
```
Accounts in this organizational directory only (Single tenant)
```

**URI de redirecionamento:**
```
Tipo: Web
URI: http://localhost:5000/auth/callback
```

Clique em **Register**

### Passo 3: Obter Client ID e Tenant ID

Após registrar, você verá a página "Overview" do aplicativo:

- **Application (client) ID** → Copie para `MICROSOFT_CLIENT_ID` no .env
- **Directory (tenant) ID** → Copie para `MICROSOFT_TENANT_ID` no .env

### Passo 4: Criar Client Secret

1. No menu lateral, clique em **Certificates & secrets**
2. Clique em **New client secret**
3. Descrição: `SharePoint AI Secret`
4. Validade: `24 months` (ou sua preferência)
5. Clique em **Add**
6. **IMPORTANTE:** Copie o **Value** IMEDIATAMENTE para `MICROSOFT_CLIENT_SECRET` no .env
   - Este valor só é mostrado uma vez!

### Passo 5: Configurar Permissões da API

1. No menu lateral, clique em **API permissions**
2. Clique em **Add a permission** → **Microsoft Graph** → **Delegated permissions**
3. Adicione as seguintes permissões:
   - `User.Read` - Ler perfil do usuário
   - `Files.Read.All` - Ler arquivos do usuário
   - `Sites.Read.All` - Ler sites do SharePoint

4. Clique em **Add permissions**
5. Clique em **Grant admin consent for [Sua Organização]**

### Passo 6: Configurar Autenticação (opcional, mas recomendado)

1. No menu lateral, clique em **Authentication**
2. Em **Implicit grant and hybrid flows**, marque:
   - ✅ ID tokens (used for implicit and hybrid flows)

## 🤖 Configuração da Claude API (Anthropic)

### Passo 1: Criar Conta na Anthropic

1. Acesse [https://console.anthropic.com/](https://console.anthropic.com/)
2. Crie uma conta ou faça login

### Passo 2: Obter API Key

1. No console, vá para **API Keys**
2. Clique em **Create Key**
3. Dê um nome: `SharePoint AI`
4. Copie a chave gerada para `CLAUDE_API_KEY` no .env

### Passo 3: Adicionar Créditos (se necessário)

- A Anthropic oferece créditos gratuitos para novos usuários
- Caso precise, adicione créditos em **Billing** → **Add credits**

## 📝 Arquivo .env Configurado

Exemplo de arquivo `.env` preenchido:

```env
# Microsoft Azure AD
MICROSOFT_CLIENT_ID=12345678-1234-1234-1234-123456789abc
MICROSOFT_CLIENT_SECRET=abc123~ABC123-ABC123abc123ABC123abc123
MICROSOFT_TENANT_ID=87654321-4321-4321-4321-cba987654321
MICROSOFT_REDIRECT_URI=http://localhost:5000/auth/callback

# SharePoint
SHAREPOINT_SITE_URL=https://fiofortei9automacaogroup.sharepoint.com/sites/Documentos

# Claude AI
CLAUDE_API_KEY=sk-ant-api03-ABC123def456GHI789jkl012MNO345pqr678STU901vwx234YZA567bcd890
CLAUDE_MODEL=claude-sonnet-4-5-20250929

# Flask
SECRET_KEY=sua-chave-secreta-aleatoria-aqui
FLASK_ENV=development
```

## ▶️ Executando a Aplicação

### 1. Ative o ambiente virtual (se não estiver ativo)

```bash
# Linux/Mac
source venv/bin/activate

# Windows
venv\Scripts\activate
```

### 2. Inicie o servidor

```bash
cd backend
python app.py
```

### 3. Acesse a aplicação

Abra seu navegador e acesse:
```
http://localhost:5000
```

## 📖 Como Usar

### 1. **Login**
- Clique em "Conectar com Microsoft"
- Faça login com sua conta Microsoft 365
- Autorize as permissões solicitadas

### 2. **Listar Documentos**
- Após o login, os documentos do SharePoint serão carregados automaticamente
- Você verá uma lista com todos os arquivos compatíveis

### 3. **Selecionar Documentos**
- Marque a caixa de seleção dos documentos que deseja resumir
- Você pode selecionar um ou múltiplos documentos

### 4. **Gerar Resumos**
- Clique em "Resumir Documentos Selecionados"
- Aguarde enquanto a IA processa os arquivos
- Os resumos serão exibidos abaixo

### 5. **Visualizar Resumos**
- Cada documento terá seu resumo individual
- Se múltiplos documentos foram selecionados, um resumo consolidado será gerado

## 📁 Estrutura do Projeto

```
sharepoint-ai-app/
├── backend/
│   ├── app.py                  # Aplicação Flask principal
│   ├── config.py               # Configurações e variáveis de ambiente
│   ├── requirements.txt        # Dependências Python
│   └── services/
│       ├── __init__.py
│       ├── microsoft_graph.py  # Serviço Microsoft Graph API
│       └── claude_ai.py        # Serviço Claude AI
├── frontend/
│   ├── index.html             # Interface principal
│   ├── css/
│   │   └── style.css          # Estilos CSS
│   └── js/
│       └── app.js             # Lógica JavaScript
├── .env.example               # Exemplo de variáveis de ambiente
├── .gitignore                 # Arquivos ignorados pelo Git
└── README.md                  # Esta documentação
```

## 🔧 Tecnologias Utilizadas

### Backend
- **Flask** - Framework web Python
- **MSAL** - Microsoft Authentication Library
- **Requests** - Cliente HTTP
- **Anthropic SDK** - Cliente Claude AI
- **PyPDF2** - Leitura de PDFs
- **python-docx** - Leitura de arquivos Word
- **openpyxl** - Leitura de arquivos Excel
- **python-pptx** - Leitura de arquivos PowerPoint

### Frontend
- **HTML5** - Estrutura
- **CSS3** - Estilização
- **JavaScript (Vanilla)** - Lógica e interações

### APIs
- **Microsoft Graph API** - Acesso ao SharePoint
- **Claude API (Anthropic)** - Processamento de linguagem natural

## 🐛 Troubleshooting

### Erro: "Configurações ausentes"

**Problema:** Variáveis de ambiente não configuradas corretamente.

**Solução:**
```bash
# Verifique se o arquivo .env existe e está preenchido
cat ../.env

# Certifique-se de que todas as variáveis estão configuradas
```

### Erro: "Não autenticado" ao listar documentos

**Problema:** Token de acesso expirado ou sessão inválida.

**Solução:**
- Faça logout e login novamente
- Verifique se as permissões do Azure AD estão corretas

### Erro ao processar documentos

**Problema:** Arquivo muito grande ou formato incompatível.

**Solução:**
- Limite de tamanho: arquivos muito grandes podem causar timeout
- Formatos suportados: PDF, DOCX, XLSX, PPTX, TXT

### Erro: "AADSTS50011" ou "AADSTS65005"

**Problema:** URI de redirecionamento não configurada corretamente no Azure.

**Solução:**
1. Vá para Azure Portal → App registrations → Seu app
2. Clique em **Authentication**
3. Verifique se `http://localhost:5000/auth/callback` está nas URIs de redirecionamento

## 🔒 Segurança

### Importante:

1. **NUNCA** compartilhe seu arquivo `.env`
2. **NUNCA** faça commit do `.env` no Git
3. Em produção, use HTTPS em vez de HTTP
4. Gere uma `SECRET_KEY` aleatória e segura
5. Revogue tokens e secrets se expostos acidentalmente

### Gerando uma SECRET_KEY segura:

```bash
python -c "import secrets; print(secrets.token_hex(32))"
```

## 📄 Licença

Este projeto é fornecido como está, para fins educacionais e de demonstração.

## 🤝 Contribuições

Contribuições são bem-vindas! Sinta-se à vontade para:

- Reportar bugs
- Sugerir novas funcionalidades
- Enviar pull requests

## 📞 Suporte

Para questões relacionadas a:

- **Azure/Microsoft Graph:** [Documentação Microsoft Graph](https://docs.microsoft.com/graph/)
- **Claude AI:** [Documentação Anthropic](https://docs.anthropic.com/)
- **Flask:** [Documentação Flask](https://flask.palletsprojects.com/)

## ✨ Autor

Desenvolvido com ❤️ para integração de SharePoint com IA.

---

**Versão:** 1.0.0
**Data:** Novembro 2024
