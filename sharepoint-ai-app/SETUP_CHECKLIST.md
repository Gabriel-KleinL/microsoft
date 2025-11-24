# ✅ Checklist de Configuração - SharePoint AI Integration

## 📋 Status Atual

### ✅ Concluído
- [x] Aplicativo registrado no Azure AD
- [x] Client ID obtido: `c70d4b2a-4933-4992-8d10-0d95f2c93e1a`
- [x] Tenant ID obtido: `b434e832-4219-4a9f-be8a-0cb8a0fae66c`
- [x] Arquivo `.env` criado com dados parciais

### ⚠️ Pendente - AÇÕES NECESSÁRIAS

## 1. 🔑 Criar Client Secret no Azure Portal

**STATUS:** ❌ NÃO FEITO

### Passos:

1. Acesse o [Azure Portal](https://portal.azure.com/)
2. Vá para **Azure Active Directory** → **App registrations**
3. Clique no aplicativo **"SharePoint AI Integration"**
4. No menu lateral, clique em **"Certificates & secrets"**
5. Clique em **"+ New client secret"**
   - Description: `SharePoint AI Secret`
   - Expires: `24 months` (recomendado)
6. Clique em **"Add"**
7. **⚠️ COPIE O VALUE IMEDIATAMENTE** (ele só aparece uma vez!)
8. Cole o valor no arquivo `.env` na variável `MICROSOFT_CLIENT_SECRET`

**Exemplo do que copiar:**
```
Value: abc123~DEF456-ghi789JKL012mno345PQR678stu901
```

---

## 2. 🔐 Configurar Permissões da API

**STATUS:** ❌ NÃO VERIFICADO

### Passos:

1. No Azure Portal, no aplicativo **"SharePoint AI Integration"**
2. No menu lateral, clique em **"API permissions"**
3. Verifique se já existe **"User.Read"** (normalmente já vem)
4. Clique em **"+ Add a permission"**
5. Selecione **"Microsoft Graph"**
6. Selecione **"Delegated permissions"**
7. Adicione estas permissões:
   - ✅ `User.Read` - Ler perfil do usuário
   - ✅ `Files.Read.All` - Ler todos os arquivos que o usuário pode acessar
   - ✅ `Sites.Read.All` - Ler sites do SharePoint

8. Clique em **"Add permissions"**
9. **IMPORTANTE:** Clique em **"Grant admin consent for [Sua Organização]"**
   - Isso autoriza as permissões para todos os usuários

---

## 3. 🌐 Verificar URI de Redirecionamento

**STATUS:** ✅ JÁ CONFIGURADO (1 Web URI)

### Verificação:

1. No Azure Portal, no aplicativo **"SharePoint AI Integration"**
2. No menu lateral, clique em **"Authentication"**
3. Verifique se existe:
   - **Platform:** Web
   - **Redirect URI:** `http://localhost:5000/auth/callback`

Se não existir, adicione:
1. Clique em **"+ Add a platform"**
2. Selecione **"Web"**
3. Em "Redirect URIs", adicione: `http://localhost:5000/auth/callback`
4. Clique em **"Configure"**

---

## 4. 🤖 Obter Claude API Key

**STATUS:** ❌ NÃO FEITO

### Passos:

1. Acesse [https://console.anthropic.com/](https://console.anthropic.com/)
2. Crie uma conta ou faça login
3. Vá para **"API Keys"** no menu lateral
4. Clique em **"Create Key"**
5. Dê um nome: `SharePoint AI`
6. Copie a chave gerada
7. Cole no arquivo `.env` na variável `CLAUDE_API_KEY`

**Exemplo:**
```
CLAUDE_API_KEY=sk-ant-api03-ABC123def456...
```

### 💰 Créditos:
- Novos usuários recebem créditos gratuitos
- Se necessário, adicione créditos em **Billing** → **Add credits**

---

## 5. 📦 Instalar Dependências Python

**STATUS:** ❌ NÃO FEITO

### Passos:

```bash
# Entre no diretório do projeto
cd sharepoint-ai-app/backend

# Crie um ambiente virtual
python3 -m venv venv

# Ative o ambiente virtual
# Linux/Mac:
source venv/bin/activate

# Windows:
venv\Scripts\activate

# Instale as dependências
pip install -r requirements.txt
```

---

## 6. ▶️ Executar a Aplicação

**STATUS:** ❌ NÃO FEITO

### Passos:

```bash
# Certifique-se de estar no diretório backend com venv ativo
cd sharepoint-ai-app/backend
source venv/bin/activate  # se não estiver ativo

# Execute a aplicação
python app.py
```

### Verificação:
- O servidor deve iniciar em: `http://localhost:5000`
- Você deve ver uma mensagem com:
  ```
  🚀 SharePoint + Claude AI Integration
  ✓ Servidor rodando em: http://localhost:5000
  ✓ SharePoint configurado: https://fiofortei9automacaogroup.sharepoint.com/sites/Documentos
  ✓ Modelo Claude: claude-3-5-sonnet-20241022
  ```

---

## 📝 Resumo dos Arquivos

### `.env` - Variáveis de Ambiente

Status atual:
```env
✅ MICROSOFT_CLIENT_ID=c70d4b2a-4933-4992-8d10-0d95f2c93e1a
❌ MICROSOFT_CLIENT_SECRET=PREENCHA_COM_SEU_CLIENT_SECRET
✅ MICROSOFT_TENANT_ID=b434e832-4219-4a9f-be8a-0cb8a0fae66c
✅ MICROSOFT_REDIRECT_URI=http://localhost:5000/auth/callback
✅ SHAREPOINT_SITE_URL=https://fiofortei9automacaogroup.sharepoint.com/sites/Documentos
❌ CLAUDE_API_KEY=PREENCHA_COM_SUA_CHAVE_CLAUDE
✅ CLAUDE_MODEL=claude-3-5-sonnet-20241022
✅ SECRET_KEY=sharepoint-ai-integration-dev-key-2024
✅ FLASK_ENV=development
```

---

## 🎯 Ordem Recomendada

1. **Criar Client Secret** (5 min) - Mais importante!
2. **Configurar Permissões API** (5 min)
3. **Obter Claude API Key** (5 min)
4. **Instalar Dependências** (2-3 min)
5. **Executar Aplicação** (1 min)
6. **Testar Login e Funcionalidades** (5 min)

**Tempo total estimado:** 20-25 minutos

---

## 🆘 Problemas Comuns

### "Configurações ausentes"
- Verifique se você preencheu `MICROSOFT_CLIENT_SECRET` e `CLAUDE_API_KEY` no arquivo `.env`

### "AADSTS50011" ou erro de redirect_uri
- Confirme que `http://localhost:5000/auth/callback` está nas URIs de redirecionamento no Azure

### "Unauthorized" ao acessar SharePoint
- Verifique se as permissões da API foram concedidas
- Clique em "Grant admin consent" no Azure Portal

### Erro ao importar módulos Python
- Certifique-se de que o ambiente virtual está ativo
- Execute `pip install -r requirements.txt` novamente

---

## ✅ Quando Tudo Estiver Pronto

Você poderá:
1. Acessar `http://localhost:5000`
2. Clicar em "Conectar com Microsoft"
3. Fazer login com sua conta Microsoft 365
4. Ver seus documentos do SharePoint
5. Selecionar documentos e gerar resumos com IA!

---

**Última atualização:** 2024-11-24
