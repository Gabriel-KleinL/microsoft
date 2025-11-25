#!/bin/bash

# Script de inicialização da aplicação SharePoint + Claude AI

echo "=================================================="
echo "🚀 SharePoint + Claude AI Integration"
echo "=================================================="

# Navega para o diretório do projeto
cd "$(dirname "$0")"

# Ativa ambiente virtual
echo "✓ Ativando ambiente virtual..."
source venv/bin/activate

# Navega para o backend
cd backend

# Inicia a aplicação
echo "✓ Iniciando servidor Flask..."
python app.py
