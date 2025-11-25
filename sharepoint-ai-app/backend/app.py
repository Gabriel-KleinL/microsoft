"""
Aplicação Flask - SharePoint + Claude AI Integration
"""
from flask import Flask, request, jsonify, redirect, session, send_from_directory
from flask_cors import CORS
from flask_caching import Cache
from config import Config
from services.microsoft_graph import MicrosoftGraphService
from services.claude_ai import ClaudeAIService
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

claude_service = ClaudeAIService(
    api_key=Config.CLAUDE_API_KEY,
    model=Config.CLAUDE_MODEL
)


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

        response_data = {
            'success': True,
            'documents': items,
            'count': len(items),
            'folders_count': len(folders),
            'files_count': len(files),
            'current_path': folder_path or '',
            'from_cache': False
        }

        # Salva no cache
        cache.set(cache_key, response_data, timeout=300)  # 5 minutos
        print(f"✓ Cache salvo para: {cache_key}")

        return jsonify(response_data)

    except Exception as e:
        print(f"Erro ao listar documentos: {e}")
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

                # Resume documento com Claude
                summary = claude_service.summarize_document(
                    file_content=file_content,
                    file_type=doc_info['type'],
                    file_name=doc_info['name']
                )

                # Salva no cache (30 minutos)
                if summary.get('success'):
                    cache.set(cache_key, summary, timeout=1800)
                    print(f"✓ Cache salvo para resumo: {doc_info['name']}")

                summary['from_cache'] = False
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
                    consolidated_summary = claude_service.generate_combined_summary(
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
