import os
import sys
import time
import requests
import urllib3
import webbrowser
from streamlit.web import cli as stcli

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Garante que o programa "abra sempre no local", ou seja, 
# force o diretório de trabalho a ser a pasta onde o .exe está rodando.
if getattr(sys, 'frozen', False):
    # Se estiver rodando compilado como .exe
    base_dir = os.path.dirname(sys.executable)
else:
    # Se estiver rodando como script normal (.py)
    base_dir = os.path.dirname(os.path.abspath(__file__))

os.chdir(base_dir)

# =====================================================================
# CONFIGURAÇÃO: Coloque aqui o link RAW do seu repositório do GitHub.
# Como pegar o link raw: Vá no arquivo no GitHub, clique em "Raw" e copie a URL da barra de endereços.
# A URL sempre começa com "https://raw.githubusercontent.com/..."
# =====================================================================
URL_APP = "https://raw.githubusercontent.com/abimaelProgammer/Gerador-de-Cardapios/main/app.py"
URL_FUNCTIONS = "https://raw.githubusercontent.com/abimaelProgammer/Gerador-de-Cardapios/main/functions.py"
URL_LOGO = "https://raw.githubusercontent.com/abimaelProgammer/Gerador-de-Cardapios/main/logo2.jpg"

FILES_TO_UPDATE = {
    "app.py": URL_APP,
    "functions.py": URL_FUNCTIONS,
    "logo2.jpg": URL_LOGO
}

def check_for_updates():
    print("Verificando se ha atualizacoes do sistema na internet...")
    for filename, url in FILES_TO_UPDATE.items():
        try:
            # Baixa com timeout maior (10 segundos) e sem verificar SSL (evita erros em proxy/exe)
            response = requests.get(url, timeout=10, verify=False)
            
            # Se a resposta for 200 OK e o arquivo não for uma página de erro
            if response.status_code == 200 and "404: Not Found" not in response.text:
                with open(filename, 'wb') as f:
                    f.write(response.content)
                print(f"[+] '{filename}' atualizado com sucesso do GitHub.")
            else:
                print(f"[-] O link para '{filename}' parece estar errado ou o repo e privado.")
        except Exception as e:
            print(f"[-] Erro ao atualizar '{filename}': {e}")
            print(f"    Usando a versao local de '{filename}'.")

if __name__ == "__main__":
    print("========================================")
    print(" GERADOR DE CARDAPIO - INICIANDO...     ")
    print("========================================")
    
    # 1. Tenta baixar a versão mais recente dos scripts
    check_for_updates()
    
    # 2. Garante que os arquivos existam localmente para rodar
    if not os.path.exists("app.py"):
        print("\n[ERRO CRITICO] O arquivo 'app.py' nao existe nesta pasta.")
        print("Tente conectar-se a internet para o sistema baixa-lo automaticamente.")
        input("Pressione ENTER para sair...")
        sys.exit(1)
        
    print("\nIniciando o servidor e abrindo o navegador...")
    time.sleep(1)

    # 3. Força a execução do Streamlit programaticamente dentro do Executável
    sys.argv = [
        "streamlit", 
        "run", 
        "app.py", 
        "--global.developmentMode=false",
        "--server.headless=true"  # Impede o Streamlit de abrir o navegador de forma nativa (dá bugs no exe)
    ]
    
    # 4. Abre o navegador manualmente na porta padrao do Streamlit
    webbrowser.open("http://localhost:8501")
    
    # 5. Entrega a execução principal para o Streamlit
    sys.exit(stcli.main())
