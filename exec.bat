#!/bin/bash

# Caminho base do projeto
cd /home/lfragoso/projetos/dash-burgetXLogComexXComercial || {
  echo "❌ Diretório não encontrado!"
  exit 1
}

# 1 - Ativar o ambiente virtual
source venv/bin/activate

# 2 - Executar main.py na pasta "mes atual"
echo "==============================="
echo "🚀 Executando main.py (mes atual)..."
echo "==============================="
python3 main.py
if [ $? -ne 0 ]; then
  echo "❌ ERRO: main.py encontrou erros."
  exit 1
fi

# 3 - Executar atualizar_systracker.py
echo "==============================="
echo "🔁 Executando atualizar_systracker.py..."
echo "==============================="
python3 atualizar_systracker.py
if [ $? -ne 0 ]; then
  echo "❌ ERRO: atualizar_systracker.py encontrou erros."
  exit 1
fi

# 4 - Executar app.py
echo "==============================="
echo "📊 Executando app.py..."
echo "==============================="
python3 app.py
if [ $? -ne 0 ]; then
  echo "❌ ERRO: app.py encontrou erros."
  exit 1
fi

echo "==============================="
echo "✅ Execução concluída com sucesso!"
echo "==============================="
