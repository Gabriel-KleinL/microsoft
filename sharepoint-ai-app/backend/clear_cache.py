#!/usr/bin/env python3
"""
Script para limpar o cache do Flask
Use este script para forçar a regeneração de resumos e testar os níveis de detalhe
"""

import sys
import os

# Adiciona o diretório backend ao path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from app import cache

def clear_cache():
    """Limpa todo o cache do Flask"""
    try:
        cache.clear()
        print("✅ Cache limpo com sucesso!")
        print("Agora você pode testar os níveis de resumo sem usar cache antigo.")
    except Exception as e:
        print(f"❌ Erro ao limpar cache: {e}")

if __name__ == "__main__":
    clear_cache()
