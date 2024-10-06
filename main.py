import os
import sys
import requests
import traceback
import msal
import mss
import io
import tkinter as tk
from tkinter import simpledialog
from datetime import datetime, timedelta
from PIL import Image, ImageDraw, ImageFont
import time
import logging

# Configurar o logging
logging.basicConfig(filename='screenshot_log.log', level=logging.INFO)

# Tenant ID e Client ID fornecidos
tenant_id = '231dd909-9b34-4b37-b58b-1f4bcc3b6ef9'  # Seu tenant ID
client_id = 'c11644f2-d053-4b65-8e95-015517ebc2d7'  # Seu client ID
sharepoint_site = "juridico"  # Nome do site SharePoint
sharepoint_tenant = "ssgruposrv.sharepoint.com"  # Substitua por seu domínio base do SharePoint
retry_limit_minutes = 30  # Tempo máximo de retry em minutos

# Lista para armazenar os caminhos temporários dos prints
imagens = []



# Função para capturar screenshots diretamente na memória e salvar localmente
def take_screenshot_all_monitors():
    with mss.mss() as sct:
        monitor = sct.monitors[0]
        img = sct.grab(monitor)

        # Converter imagem em objeto PIL
        image = Image.frombytes("RGB", img.size, img.rgb)

        # Salvar a imagem localmente
        timestamp = datetime.now().strftime("%H-%M-%S")
        file_name = f"screenshot_{timestamp}.png"
        image.save(file_name, format="PNG")

        # Armazenar o caminho do arquivo para gerar o PDF depois
        imagens.append(file_name)

        return file_name

# Função para autenticar e obter o token de acesso usando MSAL com autenticação interativa no navegador padrão
def get_access_token():
    config = {
        "authority": f"https://login.microsoftonline.com/{tenant_id}",
        "client_id": client_id,
        "scope": [f"https://{sharepoint_tenant}/.default"]
    }

    app = msal.PublicClientApplication(
        config["client_id"], authority=config["authority"]
    )

    result = app.acquire_token_interactive(scopes=config["scope"])

    if "access_token" in result:
        return result["access_token"]
    else:
        raise Exception(f"Falha na autenticação: {result.get('error_description')}")

# Função para verificar a existência de uma pasta no SharePoint
def verificar_pasta_existe(folder_path, access_token):
    try:
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Accept": "application/json;odata=verbose"
        }

        # URL para verificar a existência da pasta
        folder_url = f"https://{sharepoint_tenant}/sites/{sharepoint_site}/_api/web/GetFolderByServerRelativeUrl('{folder_path}')"

        response = requests.get(folder_url, headers=headers)

        # Se o código de resposta for 200, a pasta existe
        return response.status_code == 200

    except Exception as e:
        print(f"Erro ao verificar a pasta: {e}")
        traceback.print_exc()
        return False

# Função auxiliar para criar pastas de base, caso elas não existam
def criar_pasta_no_sharepoint_base(folder_name, parent_folder, access_token):
    try:
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json;odata=verbose"
        }

        # Remover espaços no início e no final do nome da pasta
        folder_name = folder_name.strip()

        folder_url = f"https://{sharepoint_tenant}/sites/{sharepoint_site}/_api/web/folders"
        data = {
            "__metadata": {"type": "SP.Folder"},  # Adicionando o tipo correto
            "ServerRelativeUrl": f"/sites/{sharepoint_site}/{parent_folder}/{folder_name}"
        }

        # Verificar se a pasta já existe
        folder_path = f"/sites/{sharepoint_site}/{parent_folder}/{folder_name}"
        if not verificar_pasta_existe(folder_path, access_token):
            response = requests.post(folder_url, headers=headers, json=data)
            if response.status_code == 201:
                print(f"Pasta base {folder_name} criada com sucesso no SharePoint.")
            else:
                print(f"Erro ao criar a pasta base: {response.status_code} - {response.text}")
        else:
            print(f"Pasta {folder_name} já existe no SharePoint.")

    except Exception as e:
        print(f"Erro ao criar a pasta base {folder_name}: {e}")
        traceback.print_exc()

# Função para criar a pasta no SharePoint, incluindo subpasta com o dia
def criar_pasta_no_sharepoint(folder_name, access_token):
    try:
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json;odata=verbose"
        }

        # Criar a pasta com o nome do usuário do Teams
        shared_folder_path = f"/sites/{sharepoint_site}/Documentos Compartilhados/{folder_name}"
        criar_pasta_no_sharepoint_base(folder_name, "Documentos Compartilhados", access_token)

        # Criar a subpasta com o nome do dia
        current_date = datetime.now().strftime('%Y-%m-%d')
        day_folder_name = f"{folder_name}/{current_date}"
        criar_pasta_no_sharepoint_base(current_date, f"Documentos Compartilhados/{folder_name}", access_token)

    except Exception as e:
        print("Ocorreu um erro ao tentar criar a pasta no SharePoint.")
        print(f"Erro: {e}")
        traceback.print_exc()

# Função para enviar screenshot ao SharePoint com tentativas de reenvio por 30 minutos
def enviar_screenshot_ao_sharepoint(image_path, access_token, folder_name, day_folder_name):
    try:
        with open(image_path, "rb") as img_file:
            image_bytes = img_file.read()

        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/octet-stream"
        }

        timestamp = datetime.now().strftime("%H-%M-%S")
        file_name = f"screenshot_{timestamp}.png"

        upload_url = f"https://{sharepoint_tenant}/sites/{sharepoint_site}/_api/web/GetFolderByServerRelativeUrl('/sites/{sharepoint_site}/Documentos Compartilhados/{folder_name}/{day_folder_name}')/Files/Add(url='{file_name}',overwrite=true)"

        response = requests.post(upload_url, headers=headers, data=image_bytes)

        if response.status_code == 200:
            logging.info(f"Screenshot {file_name} enviado com sucesso ao SharePoint.")
        else:
            logging.error(f"Erro ao enviar screenshot: {response.status_code} - {response.text}")

    except Exception as e:
        logging.error(f"Erro ao tentar enviar o screenshot: {e}")
        traceback.print_exc()

# Função para enviar o PDF ao SharePoint
def enviar_pdf_ao_sharepoint(pdf_path, access_token, folder_name, day_folder_name):
    try:
        with open(pdf_path, "rb") as pdf_file:
            pdf_bytes = pdf_file.read()

        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/octet-stream"
        }

        pdf_name = "screenshots.pdf"
        upload_url = f"https://{sharepoint_tenant}/sites/{sharepoint_site}/_api/web/GetFolderByServerRelativeUrl('/sites/{sharepoint_site}/Documentos Compartilhados/{folder_name}/{day_folder_name}')/Files/Add(url='{pdf_name}',overwrite=true)"

        response = requests.post(upload_url, headers=headers, data=pdf_bytes)

        if response.status_code == 200:
            print(f"PDF {pdf_name} enviado com sucesso ao SharePoint.")
        else:
            print(f"Erro ao enviar o PDF: {response.status_code} - {response.text}")

    except Exception as e:
        print(f"Erro ao tentar enviar o PDF: {e}")
        traceback.print_exc()

# Função para deletar as imagens no SharePoint
def deletar_imagens_sharepoint(access_token, folder_name, day_folder_name):
    try:
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Accept": "application/json;odata=verbose",
            "X-HTTP-Method": "DELETE",
            "If-Match": "*"
        }

        # URL para listar os arquivos no SharePoint
        list_url = f"https://{sharepoint_tenant}/sites/{sharepoint_site}/_api/web/GetFolderByServerRelativeUrl('/sites/{sharepoint_site}/Documentos Compartilhados/{folder_name}/{day_folder_name}')/Files"

        # Obter a lista de arquivos
        response = requests.get(list_url, headers=headers)

        if response.status_code == 200:
            files = response.json()["d"]["results"]
            for file in files:
                if file["Name"].endswith(".png"):
                    file_url = file["ServerRelativeUrl"]
                    # Deletar cada arquivo PNG
                    delete_url = f"https://{sharepoint_tenant}/sites/{sharepoint_site}/_api/web/GetFileByServerRelativeUrl('{file_url}')"
                    delete_response = requests.post(delete_url, headers=headers)

                    if delete_response.status_code == 200 or delete_response.status_code == 204:
                        print(f"Arquivo {file['Name']} deletado com sucesso.")
                    else:
                        print(f"Erro ao deletar {file['Name']}: {delete_response.status_code} - {delete_response.text}")
        else:
            print(f"Erro ao listar arquivos: {response.status_code} - {response.text}")

    except Exception as e:
        print(f"Erro ao deletar arquivos: {e}")
        traceback.print_exc()

# Função para gerar o PDF e excluir as imagens temporárias
def gerar_pdf_e_excluir(imagens):
    try:
        pdf_images = []
        pdf_path = "screenshots.pdf"

        # Fonte para o título (você pode ajustar o tamanho aqui)
        font = ImageFont.truetype("arial.ttf", 40)  # Exemplo com tamanho 40

        for idx, img_path in enumerate(imagens):
            img = Image.open(img_path).convert('RGB')

            # Criar uma nova imagem com margem
            width, height = img.size
            new_height = height + 100  # 100 pixels para a margem do título
            new_image = Image.new('RGB', (width, new_height), color=(255, 255, 255))
            new_image.paste(img, (0, 100))  # Desloca a imagem 100px para baixo

            # Adicionar o título na parte superior
            draw = ImageDraw.Draw(new_image)
            timestamp = datetime.now().strftime("%H:%M:%S")
            data = datetime.now().strftime("%d-%m-%Y")
            title_text = f"screenshot_{idx + 1:03d}_{data}_{timestamp}"

            # Calcular o tamanho do texto usando o novo método textbbox
            text_bbox = draw.textbbox((0, 0), title_text, font=font)
            text_width = text_bbox[2] - text_bbox[0]

            # Centralizar o texto horizontalmente
            text_x = (width - text_width) // 2
            draw.text((text_x, 10), title_text, font=font, fill=(0, 0, 0))  # Texto preto

            pdf_images.append(new_image)

        # Salvar o PDF com as imagens e seus títulos
        pdf_images[0].save(pdf_path, save_all=True, append_images=pdf_images[1:])
        print("PDF gerado com sucesso.")

        # Excluir imagens após gerar o PDF
        for img in imagens:
            os.remove(img)

        return pdf_path

    except Exception as e:
        print(f"Erro ao gerar o PDF: {e}")
        traceback.print_exc()
        return None

# Função para validar nome do Teams
def validar_nome_teams(teams_name):
    teams_name = teams_name.strip().upper()  # Remove espaços e converte para maiúsculas
    if teams_name.endswith("AGIL LTDA"):
        return teams_name
    else:
        return None

# Main
if __name__ == "__main__":
    try:
        access_token = get_access_token()

        while True:
            teams_name = simpledialog.askstring("Nome no Teams", "Digite o seu nome no Teams (deve terminar com 'AGIL LTDA'):")
            teams_name = validar_nome_teams(teams_name)

            if teams_name:
                break
            else:
                print("O nome deve terminar com 'AGIL LTDA' e estar em maiúsculas. Por favor, tente novamente.")

        criar_pasta_no_sharepoint(teams_name, access_token)

        current_date = datetime.now().strftime('%Y-%m-%d')

        # Loop para capturar e enviar 480 screenshots
        for i in range(480):
            logging.info(f"Capturando screenshot {i + 1}/480")

            screenshot_path = take_screenshot_all_monitors()

            enviar_screenshot_ao_sharepoint(screenshot_path, access_token, teams_name, current_date)

            # Aguarda 60 segundos entre as capturas
            time.sleep(60)

        # Gerar o PDF com todas as imagens capturadas e excluir os arquivos temporários
        pdf_path = gerar_pdf_e_excluir(imagens)

        # Enviar o PDF gerado para o SharePoint
        if pdf_path:
            enviar_pdf_ao_sharepoint(pdf_path, access_token, teams_name, current_date)

        # Excluir o arquivo PDF após o upload
        if os.path.exists(pdf_path):
            os.remove(pdf_path)

        # Excluir as imagens no SharePoint
        deletar_imagens_sharepoint(access_token, teams_name, current_date)

        logging.info("Processo concluído. Foram capturadas 480 screenshots, o PDF foi gerado e enviado, e as imagens foram excluídas do SharePoint.")

    except Exception as e:
        logging.error(f"Ocorreu um erro durante o processo principal: {e}")
        traceback.print_exc()
