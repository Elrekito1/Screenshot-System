import os
import requests
import shutil
import subprocess
import sys

# Configurações
APP_NAME = 'main.exe'# Nome do seu executável principal
LOCAL_VERSION_FILE = 'versao.txt' # Arquivo local que contém a versão instalada
GITHUB_REPO = 'Elrekito1/Screenshot-System'  # Repositório GitHub no formato "usuario/repo"
GITHUB_API_URL = f'https://api.github.com/repos/{GITHUB_REPO}/releases/latest'  # API para a release mais recente
DOWNLOAD_URL = 'https://github.com/Elrekito1/Screenshot-System/releases/download/v1.0.1/main.exe' # Base URL para downloads

def obter_versao_local():
    """Lê a versão local instalada."""
    try:
        with open(LOCAL_VERSION_FILE, 'r') as f:
            return f.read().strip()
    except FileNotFoundError:
        return '0.0.0'  # Se o arquivo não existir, assume que a versão é 0

def obter_versao_remota():
    """Obtém a versão mais recente disponível no GitHub Releases."""
    try:
        response = requests.get(GITHUB_API_URL)
        response.raise_for_status()
        data = response.json()
        return data['tag_name'].strip('v')  # Remove o 'v' da versão
    except requests.RequestException as e:
        print(f"Erro ao verificar versão remota: {e}")
        return None

def baixar_atualizacao(versao_remota):
    """Faz o download da nova versão do aplicativo."""
    try:
        # URL completa para o arquivo binário
        download_url = f'{DOWNLOAD_URL}v{versao_remota}/{APP_NAME}'
        response = requests.get(download_url, stream=True)
        response.raise_for_status()

        # Salva a nova versão temporariamente
        with open(f"temp_{APP_NAME}", 'wb') as f:
            shutil.copyfileobj(response.raw, f)

        # Substitui o aplicativo antigo pela nova versão
        os.remove(APP_NAME)
        shutil.move(f"temp_{APP_NAME}", APP_NAME)

        print(f"Aplicativo atualizado para a versão {versao_remota}.")
    except requests.RequestException as e:
        print(f"Erro ao baixar a atualização: {e}")
    except Exception as e:
        print(f"Erro ao substituir o aplicativo: {e}")

def iniciar_aplicacao():
    """Inicia o aplicativo atualizado."""
    try:
        subprocess.Popen([APP_NAME])
    except Exception as e:
        print(f"Erro ao iniciar {APP_NAME}: {e}")

if __name__ == '__main__':
    versao_local = obter_versao_local()
    versao_remota = obter_versao_remota()

    if versao_remota is None:
        print("Não foi possível verificar atualizações. Iniciando o aplicativo atual.")
        iniciar_aplicacao()
        sys.exit()

    if versao_local != versao_remota:
        print(f"Nova versão detectada: {versao_remota}. Atualizando...")
        baixar_atualizacao(versao_remota)
        with open(LOCAL_VERSION_FILE, 'w') as f:
            f.write(versao_remota)
    else:
        print("Você já está com a versão mais recente.")

    iniciar_aplicacao()
