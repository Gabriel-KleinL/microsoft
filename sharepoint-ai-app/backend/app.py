"""
Aplicação Flask - SharePoint + Multi-AI Integration (Claude + OpenAI)
"""
from flask import Flask, request, jsonify, redirect, session, send_from_directory
from flask_cors import CORS
from flask_caching import Cache
from config import Config
from services.microsoft_graph import MicrosoftGraphService
from services.claude_ai import ClaudeAIService
from services.openai_service import OpenAIService
import os

# Inicializa aplicação Flask
app = Flask(__name__, static_folder='../frontend')
app.secret_key = Config.SECRET_KEY
CORS(app, supports_credentials=True)

# Configura cache (5 minutos para documentos, 30 minutos para site_id)
cache = Cache(app, config={
    'CACHE_TYPE': 'SimpleCache',
    'CACHE_DEFAULT_TIMEOUT': 300  # 5 minutos
})

# Valida configurações
try:
    Config.validate()
    print("✓ Configurações validadas com sucesso")
except ValueError as e:
    print(f"✗ Erro nas configurações: {e}")
    print("  Por favor, configure o arquivo .env corretamente")

# Inicializa serviços
graph_service = MicrosoftGraphService(
    client_id=Config.MICROSOFT_CLIENT_ID,
    client_secret=Config.MICROSOFT_CLIENT_SECRET,
    tenant_id=Config.MICROSOFT_TENANT_ID
)

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


def get_ai_service():
    """Retorna o serviço de IA configurado"""
    return ai_service


# ============================================
# ROTAS DE AUTENTICAÇÃO
# ============================================

@app.route('/')
def index():
    """Serve a página principal"""
    return send_from_directory(app.static_folder, 'index.html')


@app.route('/<path:path>')
def serve_static(path):
    """Serve arquivos estáticos do frontend"""
    return send_from_directory(app.static_folder, path)


@app.route('/api/auth/login')
def login():
    """Inicia o fluxo de autenticação OAuth 2.0"""
    try:
        auth_url = graph_service.get_authorization_url(
            redirect_uri=Config.MICROSOFT_REDIRECT_URI,
            scopes=Config.SCOPES
        )
        return jsonify({'auth_url': auth_url})
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/auth/callback')
def auth_callback():
    """Callback após autenticação OAuth"""
    try:
        # Obtém código de autorização
        auth_code = request.args.get('code')
        if not auth_code:
            return redirect('/?error=no_code')

        # Troca código por token
        token_result = graph_service.acquire_token_by_auth_code(
            auth_code=auth_code,
            redirect_uri=Config.MICROSOFT_REDIRECT_URI,
            scopes=Config.SCOPES
        )

        if 'error' in token_result:
            return redirect(f'/?error={token_result["error"]}')

        # Armazena token na sessão
        session['access_token'] = token_result['access_token']

        # Obtém informações do usuário
        user_info = graph_service.get_user_info(token_result['access_token'])
        session['user_name'] = user_info.get('displayName', 'Usuário')
        session['user_email'] = user_info.get('mail', user_info.get('userPrincipalName', ''))

        return redirect('/?authenticated=true')

    except Exception as e:
        print(f"Erro no callback: {e}")
        return redirect(f'/?error={str(e)}')


@app.route('/api/auth/logout')
def logout():
    """Faz logout do usuário"""
    session.clear()
    return jsonify({'success': True})


@app.route('/api/auth/status')
def auth_status():
    """Verifica status de autenticação"""
    if 'access_token' in session:
        return jsonify({
            'authenticated': True,
            'user_name': session.get('user_name', ''),
            'user_email': session.get('user_email', '')
        })
    else:
        return jsonify({'authenticated': False})


# ============================================
# ROTAS DE SHAREPOINT
# ============================================

@app.route('/api/sharepoint/documents')
def list_documents():
    """Lista documentos e pastas do SharePoint"""
    try:
        # Verifica autenticação
        access_token = session.get('access_token')
        if not access_token:
            return jsonify({'error': 'Não autenticado'}), 401

        # Obtém parâmetro de pasta (para navegação)
        folder_path = request.args.get('folder_path', None)

        # Parâmetro para forçar refresh do cache
        force_refresh = request.args.get('force_refresh', 'false').lower() == 'true'

        # Cria chave de cache única para este usuário e pasta
        user_email = session.get('user_email', 'unknown')
        cache_key = f"docs_{user_email}_{folder_path or 'root'}"

        # Tenta obter do cache se não for refresh forçado
        if not force_refresh:
            cached_data = cache.get(cache_key)
            if cached_data:
                print(f"✓ Cache hit para: {cache_key}")
                cached_data['from_cache'] = True
                return jsonify(cached_data)

        # Obtém ID do site SharePoint (com cache de 30 minutos)
        site_cache_key = f"site_id_{Config.SHAREPOINT_SITE_URL}"
        site_id = cache.get(site_cache_key)

        if not site_id:
            site_id = graph_service.get_sharepoint_site_id(
                access_token=access_token,
                site_url=Config.SHAREPOINT_SITE_URL
            )
            cache.set(site_cache_key, site_id, timeout=1800)  # 30 minutos

        # Lista documentos e pastas
        items = graph_service.list_documents(
            access_token=access_token,
            site_id=site_id,
            folder_path=folder_path
        )

        # Separa pastas e arquivos para estatísticas
        folders = [item for item in items if item.get('isFolder')]
        files = [item for item in items if not item.get('isFolder')]

        # Obtém informações da pasta atual (se estiver navegando em uma pasta)
        folder_name = None
        web_url = Config.SHAREPOINT_SITE_URL

        if folder_path and items and len(folders) > 0:
            # Tenta obter o nome da biblioteca/pasta do primeiro item
            folder_name = folders[0].get('driveName', 'Biblioteca')

        response_data = {
            'success': True,
            'documents': items,
            'count': len(items),
            'folders_count': len(folders),
            'files_count': len(files),
            'current_path': folder_path or '',
            'folder_name': folder_name,
            'web_url': web_url,
            'from_cache': False
        }

        # Salva no cache
        cache.set(cache_key, response_data, timeout=300)  # 5 minutos
        print(f"✓ Cache salvo para: {cache_key}")

        return jsonify(response_data)

    except Exception as e:
        print(f"❌ ERRO ao listar documentos: {e}")
        import traceback
        traceback.print_exc()

        # Se erro 401, limpa sessão para forçar novo login
        if '401' in str(e) or 'Unauthorized' in str(e):
            session.clear()
            return jsonify({'error': 'Token expirado. Faça login novamente.', 'auth_required': True}), 401

        return jsonify({'error': str(e)}), 500


@app.route('/api/sharepoint/search')
def search_documents():
    """Busca documentos em todo o SharePoint"""
    try:
        # Verifica autenticação
        access_token = session.get('access_token')
        if not access_token:
            return jsonify({'error': 'Não autenticado'}), 401

        # Obtém query de busca
        query = request.args.get('q', '').strip()
        if not query:
            return jsonify({'error': 'Query de busca não fornecida'}), 400

        # Limite de resultados
        max_results = int(request.args.get('limit', 200))

        # Cria chave de cache única para esta busca
        user_email = session.get('user_email', 'unknown')
        cache_key = f"search_{user_email}_{query}_{max_results}"

        # Tenta obter do cache
        cached_data = cache.get(cache_key)
        if cached_data:
            print(f"✓ Cache hit para busca: {query}")
            cached_data['from_cache'] = True
            return jsonify(cached_data)

        # Obtém ID do site SharePoint (com cache)
        site_cache_key = f"site_id_{Config.SHAREPOINT_SITE_URL}"
        site_id = cache.get(site_cache_key)

        if not site_id:
            site_id = graph_service.get_sharepoint_site_id(
                access_token=access_token,
                site_url=Config.SHAREPOINT_SITE_URL
            )
            cache.set(site_cache_key, site_id, timeout=1800)  # 30 minutos

        # Realiza busca global
        results = graph_service.search_all_documents(
            access_token=access_token,
            site_id=site_id,
            query=query,
            max_results=max_results
        )

        response_data = {
            'success': True,
            'documents': results,
            'count': len(results),
            'query': query,
            'from_cache': False
        }

        # Salva no cache (5 minutos)
        cache.set(cache_key, response_data, timeout=300)
        print(f"✓ Cache salvo para busca: {query}")

        return jsonify(response_data)

    except Exception as e:
        print(f"Erro ao buscar documentos: {e}")
        return jsonify({'error': str(e)}), 500


# ============================================
# ROTAS DE IA / RESUMO
# ============================================

@app.route('/api/ai/summarize', methods=['POST'])
def summarize_documents():
    """Resume um ou múltiplos documentos usando Claude AI"""
    try:
        # Verifica autenticação
        access_token = session.get('access_token')
        if not access_token:
            return jsonify({'error': 'Não autenticado'}), 401

        # Obtém documentos selecionados
        data = request.json
        document_ids = data.get('document_ids', [])

        if not document_ids:
            return jsonify({'error': 'Nenhum documento selecionado'}), 400

        # Processa cada documento
        summaries = []
        cache_hits = 0

        for doc_info in document_ids:
            try:
                # Cria chave de cache para este documento
                cache_key = f"summary_{doc_info['id']}_{doc_info['driveId']}"

                # Tenta obter resumo do cache
                cached_summary = cache.get(cache_key)
                if cached_summary:
                    print(f"✓ Cache hit para resumo: {doc_info['name']}")
                    cached_summary['from_cache'] = True
                    summaries.append(cached_summary)
                    cache_hits += 1
                    continue

                # Baixa conteúdo do arquivo
                file_content = graph_service.download_file_content(
                    access_token=access_token,
                    drive_id=doc_info['driveId'],
                    file_id=doc_info['id']
                )

                # Resume documento com IA
                summary = get_ai_service().summarize_document(
                    file_content=file_content,
                    file_type=doc_info['type'],
                    file_name=doc_info['name']
                )

                # Salva no cache (30 minutos)
                if summary.get('success'):
                    summary['document_id'] = doc_info['id']  # Adiciona ID do documento
                    cache.set(cache_key, summary, timeout=1800)
                    print(f"✓ Cache salvo para resumo: {doc_info['name']}")

                summary['from_cache'] = False
                summary['document_id'] = doc_info['id']  # Garante que sempre tenha o ID
                summaries.append(summary)

            except Exception as e:
                summaries.append({
                    'success': False,
                    'error': str(e),
                    'file_name': doc_info.get('name', 'Desconhecido'),
                    'from_cache': False
                })

        # Se múltiplos documentos, gera resumo consolidado
        consolidated_summary = None
        if len(summaries) > 1:
            successful_summaries = [s for s in summaries if s.get('success')]
            if successful_summaries:
                summary_texts = [s['summary'] for s in successful_summaries]
                file_names = [s['file_name'] for s in successful_summaries]

                # Cache do resumo consolidado
                doc_ids_str = '_'.join(sorted([d['id'] for d in document_ids]))
                consolidated_cache_key = f"consolidated_{doc_ids_str}"

                consolidated_summary = cache.get(consolidated_cache_key)
                if not consolidated_summary:
                    consolidated_summary = get_ai_service().generate_combined_summary(
                        individual_summaries=summary_texts,
                        file_names=file_names
                    )
                    cache.set(consolidated_cache_key, consolidated_summary, timeout=1800)
                    print("✓ Cache salvo para resumo consolidado")
                else:
                    print("✓ Cache hit para resumo consolidado")

        return jsonify({
            'success': True,
            'summaries': summaries,
            'consolidated_summary': consolidated_summary,
            'total_processed': len(summaries),
            'cache_hits': cache_hits
        })

    except Exception as e:
        print(f"Erro ao resumir documentos: {e}")
        return jsonify({'error': str(e)}), 500


@app.route('/api/cache/clear', methods=['POST'])
def clear_cache():
    """Limpa o cache da aplicação"""
    try:
        cache.clear()
        return jsonify({
            'success': True,
            'message': 'Cache limpo com sucesso'
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/ai/chat', methods=['POST'])
def chat():
    """Chat com IA usando RAG (Retrieval Augmented Generation)"""
    try:
        # Verifica autenticação
        access_token = session.get('access_token')
        if not access_token:
            return jsonify({'error': 'Não autenticado'}), 401

        # Obtém dados da requisição
        data = request.get_json()
        user_message = data.get('message', '').strip()
        conversation_history = data.get('conversation_history', [])

        if not user_message:
            return jsonify({'error': 'Mensagem vazia'}), 400

        # Log header
        print(f"\n{'='*70}")
        print(f"║ 🤖 SOFIA - PROCESSAMENTO DE CHAT")
        print(f"{'='*70}")
        print(f"👤 Usuário: {session.get('user_email', 'unknown')}")
        print(f"📝 Pergunta: {user_message}")
        print(f"⏰ Timestamp: {__import__('datetime').datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"💬 Histórico: {len(conversation_history)} mensagens anteriores")
        print(f"{'='*70}\n")

        # Obtém ID do site SharePoint (com cache)
        site_cache_key = f"site_id_{Config.SHAREPOINT_SITE_URL}"
        site_id = cache.get(site_cache_key)

        if not site_id:
            print("🔐 Obtendo site_id do SharePoint...")
            site_id = graph_service.get_sharepoint_site_id(
                access_token=access_token,
                site_url=Config.SHAREPOINT_SITE_URL
            )
            cache.set(site_cache_key, site_id, timeout=1800)
            print(f"✓ Site ID obtido: {site_id[:20]}...")
        else:
            print(f"✓ Site ID (cache): {site_id[:20]}...")

        # Etapa 1: Buscar documentos relevantes usando busca global
        print(f"\n{'─'*70}")
        print("🔍 ETAPA 1: BUSCANDO DOCUMENTOS NO SHAREPOINT")
        print(f"{'─'*70}")
        print(f"📊 Query de busca: '{user_message}'")
        print(f"📈 Máximo de resultados: 10")

        relevant_docs = graph_service.search_all_documents(
            access_token=access_token,
            site_id=site_id,
            query=user_message,
            max_results=10
        )

        print(f"\n✅ Busca concluída: {len(relevant_docs)} documentos encontrados")
        if relevant_docs:
            print("📄 Documentos encontrados:")
            for i, doc in enumerate(relevant_docs[:5], 1):
                print(f"   {i}. {doc['name']} ({doc['type']}) - Drive: {doc.get('driveName', 'N/A')}")

        # Etapa 2: Extrair conteúdo dos documentos (máximo 3 documentos)
        print(f"\n{'─'*70}")
        print("📄 ETAPA 2: EXTRAINDO CONTEÚDO DOS DOCUMENTOS")
        print(f"{'─'*70}")

        documents_content = []
        max_docs_to_analyze = 3
        print(f"📊 Documentos a analisar: {min(len(relevant_docs), max_docs_to_analyze)}")

        for i, doc in enumerate(relevant_docs[:max_docs_to_analyze], 1):
            try:
                print(f"\n   [{i}/{max_docs_to_analyze}] 📄 {doc['name']}")
                print(f"       Tipo: {doc['type']}")
                print(f"       Tamanho: {doc.get('size', 0)} bytes")

                # Baixa conteúdo do arquivo
                print(f"       ⬇️  Baixando arquivo...")
                file_content = graph_service.download_file_content(
                    access_token=access_token,
                    drive_id=doc['driveId'],
                    file_id=doc['id']
                )
                print(f"       ✓ Download concluído ({len(file_content)} bytes)")

                # Extrai texto
                print(f"       📝 Extraindo texto...")
                extracted_text = get_ai_service().extract_text_from_file(
                    file_content=file_content,
                    file_type=doc['type']
                )

                # Limita tamanho do texto
                max_chars = 10000
                was_truncated = False
                if len(extracted_text) > max_chars:
                    extracted_text = extracted_text[:max_chars] + "\n[... conteúdo truncado ...]"
                    was_truncated = True

                documents_content.append({
                    'name': doc['name'],
                    'type': doc['type'],
                    'content': extracted_text,
                    'webUrl': doc.get('webUrl', ''),
                    'id': doc['id'],
                    'driveId': doc['driveId']
                })

                truncated_msg = " (truncado)" if was_truncated else ""
                print(f"       ✅ Texto extraído: {len(extracted_text):,} caracteres{truncated_msg}")
                print(f"       📊 Preview: {extracted_text[:100].strip()}...")

            except Exception as e:
                print(f"       ❌ ERRO: {str(e)}")
                import traceback
                print(f"       Stack: {traceback.format_exc()[:200]}...")
                continue

        print(f"\n✅ Extração concluída: {len(documents_content)} documentos processados com sucesso")

        # Etapa 3: Construir contexto para IA
        print(f"\n{'─'*70}")
        print(f"🧠 ETAPA 3: CONSTRUINDO CONTEXTO PARA {ai_provider_name.upper()} AI")
        print(f"{'─'*70}")

        context_text = ""
        if documents_content:
            context_text = "\n\n".join([
                f"DOCUMENTO: {doc['name']}\n{'='*50}\n{doc['content']}\n"
                for doc in documents_content
            ])
            total_context_chars = len(context_text)
            print(f"📊 Contexto construído: {total_context_chars:,} caracteres")
            print(f"📄 Documentos no contexto: {len(documents_content)}")
            for i, doc in enumerate(documents_content, 1):
                print(f"   {i}. {doc['name']} - {len(doc['content']):,} chars")
        else:
            print("⚠️  Nenhum documento com conteúdo válido")

        # Etapa 4: Criar prompt para IA com RAG
        print(f"\n{'─'*70}")
        print(f"✨ ETAPA 4: GERANDO RESPOSTA COM {ai_provider_name.upper()} AI")
        print(f"{'─'*70}")

        system_prompt = f"""Você é Sofia, uma assistente IA especializada em ajudar usuários a encontrar informações no SharePoint.

CONTEXTO DOS DOCUMENTOS:
{context_text if context_text else "Nenhum documento relevante encontrado."}

INSTRUÇÕES:
- Responda a pergunta do usuário baseando-se APENAS nas informações dos documentos fornecidos
- Se a informação não estiver nos documentos, diga claramente que não encontrou
- Cite os nomes dos documentos quando usar informações deles
- Seja objetivo, prestativa e direto ao ponto
- Use formatação markdown quando apropriado (**negrito**, listas, etc.)
- Mantenha um tom profissional mas amigável"""

        # Constrói histórico de conversação
        messages = []

        # Adiciona mensagens anteriores (se houver)
        for msg in conversation_history[-4:]:
            role = "user" if msg['sender'] == 'user' else "assistant"
            messages.append({
                "role": role,
                "content": msg['content']
            })

        # Adiciona mensagem atual
        messages.append({
            "role": "user",
            "content": user_message
        })

        print(f"📤 Parâmetros da requisição:")
        print(f"   Modelo: {get_ai_service().model}")
        print(f"   Max tokens: 2000")
        print(f"   System prompt: {len(system_prompt):,} caracteres")
        print(f"   Mensagens no histórico: {len(messages)}")
        print(f"   Pergunta: \"{user_message}\"")
        print(f"\n⏳ Aguardando resposta de {ai_provider_name.upper()} AI...")

        # Chama IA (funciona para Claude e OpenAI)
        import time
        start_time = time.time()

        if ai_provider_name == 'claude':
            # Claude usa client.messages.create
            response = ai_service.client.messages.create(
                model=ai_service.model,
                max_tokens=2000,
                system=system_prompt,
                messages=messages
            )
            ai_response = response.content[0].text
            input_tokens = response.usage.input_tokens
            output_tokens = response.usage.output_tokens
        else:  # openai
            # OpenAI usa answer_question
            result = ai_service.answer_question(
                question=user_message,
                documents_context=context_text,
                conversation_history=conversation_history
            )
            ai_response = result['response']
            input_tokens = 0  # OpenAI não retorna tokens facilmente
            output_tokens = 0

        elapsed_time = time.time() - start_time

        print(f"✅ Resposta recebida!")
        print(f"   ⏱️  Tempo de resposta: {elapsed_time:.2f}s")
        print(f"   📝 Tamanho da resposta: {len(ai_response)} caracteres")
        if input_tokens > 0:
            print(f"   📊 Tokens usados: ~{input_tokens} input / ~{output_tokens} output")
        print(f"   💬 Preview: {ai_response[:150].strip()}...")

        print(f"\n{'='*70}")
        print(f"✅ PROCESSAMENTO COMPLETO - SOFIA RESPONDEU!")
        print(f"{'='*70}\n")

        # Prepara fontes para retornar
        sources = [
            {
                'name': doc['name'],
                'type': doc['type'],
                'webUrl': doc['webUrl']
            }
            for doc in documents_content
        ]

        return jsonify({
            'success': True,
            'response': ai_response,
            'sources': sources,
            'documents_analyzed': len(documents_content),
            'documents_found': len(relevant_docs)
        })

    except Exception as e:
        print(f"❌ Erro no chat: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/api/health')
def health_check():
    """Verifica saúde da aplicação"""
    return jsonify({
        'status': 'healthy',
        'services': {
            'microsoft_graph': bool(Config.MICROSOFT_CLIENT_ID),
            'claude_ai': bool(Config.CLAUDE_API_KEY),
            'sharepoint': bool(Config.SHAREPOINT_SITE_URL)
        }
    })


# ============================================
# INICIALIZAÇÃO
# ============================================

if __name__ == '__main__':
    print("\n" + "="*50)
    print("🚀 SharePoint + Claude AI Integration")
    print("="*50)
    print(f"✓ Servidor rodando em: http://localhost:5000")
    print(f"✓ SharePoint configurado: {Config.SHAREPOINT_SITE_URL}")
    print(f"✓ Modelo Claude: {Config.CLAUDE_MODEL}")
    print("="*50 + "\n")

    app.run(
        debug=(Config.FLASK_ENV == 'development'),
        host='0.0.0.0',
        port=5000
    )
