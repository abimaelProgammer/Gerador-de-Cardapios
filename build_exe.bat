@echo off
echo ============================================================
echo PREPARANDO O EXECUTAVEL DO GERADOR DE CARDAPIO...
echo ============================================================
echo.

echo 1. Instalando o PyInstaller...
.venv\Scripts\python.exe -m pip install pyinstaller

echo.
echo 2. Criando o executavel... (Isso pode demorar alguns minutos)
echo.

:: O comando abaixo pega o arquivo run_app.py e transforma em um .exe.
:: --onefile: Junta tudo em um unico arquivo .exe
:: --collect-all: Forca o PyInstaller a empacotar pastas escondidas de certas bibliotecas
:: --icon=NONE (opcional, pode adicionar um ícone depois)
.venv\Scripts\python.exe -m PyInstaller --onefile --noconfirm --collect-all streamlit --collect-all altair --collect-data openpyxl --collect-data requests -n GeradorCardapio run_app.py

echo.
echo ============================================================
echo CONCLUIDO!
echo O seu executavel estara disponivel na pasta: "dist\GeradorCardapio.exe"
echo ============================================================
pause
