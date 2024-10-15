import ctypes
import os
import sys
import requests
import traceback
import msal
import mss
import io
import tkinter as tk
from tkinter import simpledialog, messagebox
from datetime import datetime, timedelta
from PIL import Image, ImageDraw, ImageFont
import time
import logging
import signal
from PyPDF2 import PdfReader, PdfWriter
import json
import threading
import atexit
import shutil
import re

# Função para exibir o termo de compromisso e obter a aceitação do usuário
def exibir_termo():
    termo = (
        "Termo de Compromisso\n\n"
        "Ao utilizar este programa, você concorda com as seguintes condições:\n"
        "1. Este software captura screenshots do seu computador periodicamente e os envia para um servidor remoto.\n"
        "2. Todos os dados capturados são tratados de acordo com as diretrizes da LGPD (Lei Geral de Proteção de Dados).\n"
        "3. Nenhuma informação pessoal sensível será coletada sem o seu consentimento explícito.\n"
        "4. Você é responsável por garantir que os dados compartilhados através do programa não contenham informações sigilosas que não devam ser compartilhadas.\n"
        "5. O uso do programa está de acordo com as políticas de privacidade e segurança da sua organização.\n"
        "6. Você pode encerrar o uso deste software a qualquer momento, mas o uso continuado implica na aceitação dos termos.\n"
        "\nSe você não aceitar os termos, o programa será encerrado."
    )

    return messagebox.askyesno("Termo de Compromisso", termo)

def verificar_aceitacao_termo():
    termo_file = "aceitacao_termo.txt"
    if os.path.exists(termo_file):
        with open(termo_file, "r") as f:
            return f.read().strip() == "aceito"
    else:
        return False

def salvar_aceitacao_termo():
    with open("aceitacao_termo.txt", "w") as f:
        f.write("aceito")

# Ocultar a janela do console (somente no Windows)
def hide_console():
    if os.name == 'nt':  # Somente no Windows
        ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 0)

# Chamar a função para ocultar o console
hide_console()

# Configuração do logging
logging.basicConfig(filename='screenshot_log.log', level=logging.INFO)

# Tenant ID e Client ID fornecidos
tenant_id = '231dd909-9b34-4b37-b58b-1f4bcc3b6ef9'
client_id = 'c11644f2-d053-4b65-8e95-015517ebc2d7'
sharepoint_site = "juridico"
sharepoint_tenant = "ssgruposrv.sharepoint.com"
retry_limit_minutes = 60

# Lista para armazenar os caminhos temporários dos prints
imagens = []

# Caminho para a pasta oculta onde as imagens serão salvas
hidden_folder = os.path.join(os.getcwd(), ".hidden_screenshots")

# Função para capturar screenshots e salvar localmente na pasta oculta
def take_screenshot_all_monitors():
    with mss.mss() as sct:
        monitor = sct.monitors[0]
        img = sct.grab(monitor)

        # Converter imagem em objeto PIL
        image = Image.frombytes("RGB", img.size, img.rgb)

        # Salvar a imagem na pasta oculta
        timestamp = datetime.now().strftime("%H-%M-%S")
        file_name = f"screenshot_{timestamp}.png"
        file_path = os.path.join(hidden_folder, file_name)
        image.save(file_path, format="PNG")

        # Armazenar o caminho do arquivo para gerar o PDF depois
        imagens.append(file_path)

        return file_path

# Função para autenticar e obter o token de acesso usando MSAL com autenticação interativa
def get_access_token():
    config = {
        "authority": f"https://login.microsoftonline.com/{tenant_id}",
        "client_id": client_id,
        "scope": [f"https://{sharepoint_tenant}/.default"]
    }

    app = msal.PublicClientApplication(config["client_id"], authority=config["authority"])
    result = app.acquire_token_interactive(scopes=config["scope"])

    if "access_token" in result:
        return result["access_token"]
    else:
        raise Exception(f"Falha na autenticação: {result.get('error_description')}")



# Função para criar o ambiente no SharePoint (pastas, se necessário)
def criar_ambiente_sharepoint(folder_name, day_folder_name, access_token):
    try:
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Accept": "application/json;odata=verbose",
            "Content-Type": "application/json;odata=verbose"
        }

        folder_url = f"https://{sharepoint_tenant}/sites/{sharepoint_site}/_api/web/folders"
        data = {
            "__metadata": {"type": "SP.Folder"},
            "ServerRelativeUrl": f"/sites/{sharepoint_site}/Documentos Compartilhados/{folder_name}"
        }

        response = requests.post(folder_url, headers=headers, data=json.dumps(data))

        if response.status_code == 201:
            print(f"Pasta {folder_name} criada com sucesso.")
        else:
            print(f"Pasta {folder_name} já existe ou houve um erro. Status: {response.status_code}")

        # Criar pasta por dia, se necessário
        data_day = {
            "__metadata": {"type": "SP.Folder"},
            "ServerRelativeUrl": f"/sites/{sharepoint_site}/Documentos Compartilhados/{folder_name}/{day_folder_name}"
        }
        response_day = requests.post(folder_url, headers=headers, json=data_day)
        if response_day.status_code == 201:
            print(f"Pasta {day_folder_name} criada com sucesso.")
        else:
            print(f"Pasta {day_folder_name} já existe ou houve um erro. Status: {response_day.status_code}")

    except Exception as e:
        print(f"Erro ao criar ambiente no SharePoint: {e}")
        traceback.print_exc()

# Função para enviar screenshots ao SharePoint
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

# Função para enviar o PDF ao SharePoint e deletá-lo localmente, com retry em caso de erro
def enviar_pdf_e_excluir_local(pdf_path, access_token, folder_name, day_folder_name):
    start_time = datetime.now()

    while (datetime.now() - start_time) < timedelta(minutes=retry_limit_minutes):
        try:
            with open(pdf_path, "rb") as pdf_file:
                pdf_bytes = pdf_file.read()

            headers = {
                "Authorization": f"Bearer {access_token}",
                "Content-Type": "application/octet-stream"
            }

            timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
            pdf_name = f"screenshots_{timestamp}.pdf"
            upload_url = f"https://{sharepoint_tenant}/sites/{sharepoint_site}/_api/web/GetFolderByServerRelativeUrl('/sites/{sharepoint_site}/Documentos Compartilhados/{folder_name}/{day_folder_name}')/Files/Add(url='{pdf_name}',overwrite=false)"

            response = requests.post(upload_url, headers=headers, data=pdf_bytes)

            if response.status_code == 200:
                print(f"PDF {pdf_name} enviado com sucesso ao SharePoint.")
                if os.path.exists(pdf_path):
                    os.remove(pdf_path)
                    print(f"PDF {pdf_name} foi excluído localmente.")
                return

            else:
                print(f"Erro ao enviar o PDF: {response.status_code} - {response.text}")

        except Exception as e:
            print(f"Erro ao tentar enviar o PDF: {e}")
            traceback.print_exc()

        print("Tentando novamente em 60 segundos...")
        time.sleep(60)

    print("Não foi possível enviar o PDF após 1 hora de tentativas.")

# Função para deletar os arquivos PNG do SharePoint com retry em caso de erro
def deletar_pngs_do_sharepoint(folder_name, day_folder_name, access_token):
    start_time = datetime.now()

    while (datetime.now() - start_time) < timedelta(minutes=retry_limit_minutes):
        try:
            headers = {
                "Authorization": f"Bearer {access_token}",
                "Accept": "application/json;odata=verbose"
            }

            list_files_url = f"https://{sharepoint_tenant}/sites/{sharepoint_site}/_api/web/GetFolderByServerRelativeUrl('/sites/{sharepoint_site}/Documentos Compartilhados/{folder_name}/{day_folder_name}')/Files"

            response = requests.get(list_files_url, headers=headers)

            if response.status_code == 200:
                files = response.json()["d"]["results"]
                for file in files:
                    if file["Name"].endswith(".png"):
                        file_url = file["ServerRelativeUrl"]
                        delete_url = f"https://{sharepoint_tenant}/sites/{sharepoint_site}/_api/web/GetFileByServerRelativeUrl('{file_url}')"

                        delete_response = requests.delete(delete_url, headers=headers)
                        if delete_response.status_code == 200:
                            print(f"PNG {file['Name']} deletado com sucesso do SharePoint.")
                        else:
                            print(f"Erro ao deletar PNG {file['Name']} do SharePoint: {delete_response.status_code} - {delete_response.text}")
                return

            else:
                print(f"Erro ao listar arquivos: {response.status_code} - {response.text}")

        except Exception as e:
            print(f"Erro ao deletar PNGs: {e}")
            traceback.print_exc()

        print("Tentando novamente em 60 segundos...")
        time.sleep(60)

    print("Não foi possível deletar os PNGs após 1 hora de tentativas.")

# Função para gerar o PDF
def gerar_pdf_e_excluir(imagens):
    try:
        pdf_writer = PdfWriter()
        pdf_path = "screenshots.pdf"

        font = ImageFont.truetype("arial.ttf", 40)

        for idx, img_path in enumerate(imagens):
            img = Image.open(img_path).convert('RGB')
            width, height = img.size
            new_height = height + 100
            new_image = Image.new('RGB', (width, new_height), color=(255, 255, 255))
            new_image.paste(img, (0, 100))

            draw = ImageDraw.Draw(new_image)
            timestamp = datetime.now().strftime("%H:%M:%S")
            data = datetime.now().strftime("%d-%m-%Y")
            title_text = f"screenshot_{idx + 1:03d}_{data}_{timestamp}"

            text_bbox = draw.textbbox((0, 0), title_text, font=font)
            text_width = text_bbox[2] - text_bbox[0]
            text_x = (width - text_width) // 2
            draw.text((text_x, 10), title_text, font=font, fill=(0, 0, 0))

            img_pdf_path = img_path.replace(".png", ".pdf")
            new_image.save(img_pdf_path, "PDF", resolution=100)

            img_reader = PdfReader(img_pdf_path)
            for page in img_reader.pages:
                pdf_writer.add_page(page)

        with open(pdf_path, "wb") as output_pdf:
            pdf_writer.write(output_pdf)

        print("PDF gerado com sucesso.")

        for img in imagens:
            os.remove(img)

        return pdf_path

    except Exception as e:
        print(f"Erro ao gerar o PDF: {e}")
        traceback.print_exc()
        return None

# Função para excluir e recriar a pasta oculta com verificação detalhada
def recriar_pasta_oculta():
    try:
        # Verifica se a pasta existe
        if os.path.exists(hidden_folder):
            print(f"Tentando excluir a pasta {hidden_folder}...")
            # Remove toda a pasta e seu conteúdo
            try:
                shutil.rmtree(hidden_folder)
                print(f"Pasta {hidden_folder} e todo o seu conteúdo foram excluídos com sucesso.")
            except Exception as e:
                print(f"Erro ao tentar excluir a pasta: {e}")
                traceback.print_exc()

        # Após tentar excluir, verificar se a pasta realmente foi removida
        if os.path.exists(hidden_folder):
            print(f"ERRO: A pasta {hidden_folder} ainda existe após tentativa de exclusão.")
        else:
            print(f"Pasta {hidden_folder} removida com sucesso.")

        # Recriar a pasta vazia
        os.makedirs(hidden_folder)
        print(f"Pasta {hidden_folder} recriada com sucesso.")

    except Exception as e:
        print(f"Erro ao excluir/recriar a pasta oculta: {e}")
        traceback.print_exc()

# Função para capturar o sinal de término (Ctrl+C ou encerramento do programa)
def signal_handler(sig, frame):
    global access_token, teams_name, current_date
    fechar_programa_e_gerar_pdf(access_token, teams_name, current_date)
    # Ao fechar, recriar a pasta para garantir que ela seja limpa
    recriar_pasta_oculta()
    sys.exit(0)

# Função para capturar e enviar os prints (Thread Separada)
def capturar_e_enviar(access_token, teams_name, current_date):
    i = 0
    while True:
        logging.info(f"Capturando screenshot {i + 1}")
        screenshot_path = take_screenshot_all_monitors()
        enviar_screenshot_ao_sharepoint(screenshot_path, access_token, teams_name, current_date)
        time.sleep(60)
        i += 1

# Função para finalizar o programa e gerar o PDF
def fechar_programa_e_gerar_pdf(access_token, teams_name, current_date):
    print("\nEncerrando o programa... Gerando PDF e excluindo arquivos temporários.")
    try:
        if imagens:
            pdf_path = gerar_pdf_e_excluir(imagens)
            if pdf_path:
                enviar_pdf_e_excluir_local(pdf_path, access_token, teams_name, current_date)
                deletar_pngs_do_sharepoint(teams_name, current_date, access_token)
            else:
                logging.error("Erro ao gerar o PDF. Nenhum PDF foi gerado.")
        else:
            logging.info("Nenhuma imagem disponível para gerar PDF.")
    except Exception as e:
        logging.error(f"Erro durante o processo de fechamento: {e}")
        traceback.print_exc()

# Função para lidar com a finalização da aplicação
def on_closing():
    global access_token, teams_name, current_date
    if messagebox.askokcancel("Encerrar", "Deseja realmente encerrar o programa?"):
        try:
            print("Iniciando o processo de fechamento...")
            fechar_programa_e_gerar_pdf(access_token, teams_name, current_date)
            recriar_pasta_oculta()  # Exclui a pasta ao fechar o programa
        except Exception as e:
            logging.error(f"Erro ao tentar fechar o programa corretamente: {e}")
            traceback.print_exc()
        finally:
            print("Fechamento completo. Encerrando o programa.")
            root.destroy()

# Função que roda a captura de prints em um thread separado
def iniciar_captura_em_thread(access_token, teams_name, current_date):
    captura_thread = threading.Thread(target=capturar_e_enviar, args=(access_token, teams_name, current_date), daemon=True)
    captura_thread.start()


# Função para validar nome do Teams
def validar_nome_teams(teams_name):
    teams_name = teams_name.strip().upper()

    # Expressão regular para verificar se existe um número separado do nome
    pattern_com_numero = r"([A-Z]+)\s?(\d+)\s?AGIL LTDA$"

    # Se houver um número separado do nome, corrige para juntar o número ao nome
    match = re.match(pattern_com_numero, teams_name)
    if match:
        nome_base = match.group(1)  # Parte do nome (ex: COMERCIAL, JURIDICO)
        numero = match.group(2)     # Número
        return f"{nome_base}{numero} AGIL LTDA"

    # Se não houver número no nome, retorna o nome como está
    pattern_sem_numero = r"^[A-Z]+ AGIL LTDA$"
    if re.match(pattern_sem_numero, teams_name):
        return teams_name

    # Se não seguir nenhum dos padrões, retorna None (inválido)
    return None
# Função para listar pastas no SharePoint
def listar_pastas_sharepoint(access_token):
    try:
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Accept": "application/json"
        }

        # URL para listar as pastas dentro da biblioteca 'Documentos Compartilhados'
        list_folders_url = f"https://{sharepoint_tenant}/sites/{sharepoint_site}/_api/web/GetFolderByServerRelativeUrl('/sites/{sharepoint_site}/Documentos Compartilhados')/Folders"

        response = requests.get(list_folders_url, headers=headers)

        if response.status_code == 200:
            # Retornar os nomes das pastas
            folders = response.json()["d"]["results"]
            return [folder["Name"] for folder in folders]
        else:
            print(f"Erro ao listar pastas: {response.status_code} - {response.text}")
            return []

    except Exception as e:
        print(f"Erro ao listar pastas no SharePoint: {e}")
        traceback.print_exc()
        return []
# Função para verificar e renomear pastas com espaços no nome
def verificar_e_corrigir_pastas_incorretas(access_token):
    pastas_incorretas = listar_pastas_sharepoint(access_token)  # Listar pastas existentes

    for pasta in pastas_incorretas:
        # Verifica se o nome da pasta contém "comercial " seguido por um número separado (com espaço)
        if "COMERCIAL " in pasta:
            # Corrige o nome da pasta removendo o espaço extra
            nome_corrigido = pasta.replace("COMERCIAL ", "COMERCIAL")

            print(f"Renomeando a pasta {pasta} para {nome_corrigido}")
            renomear_pasta(access_token, pasta, nome_corrigido)

# Função para renomear uma pasta no SharePoint
def renomear_pasta(access_token, old_folder_name, new_folder_name):
    try:
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Accept": "application/json;odata=verbose",
            "Content-Type": "application/json;odata=verbose"
        }

        # URL da API para renomear a pasta
        url = f"https://{sharepoint_tenant}/sites/{sharepoint_site}/_api/web/GetFolderByServerRelativeUrl('/sites/{sharepoint_site}/Documentos Compartilhados/{old_folder_name}')/ListItemAllFields"

        # Dados para alterar o nome da pasta
        data = {
            "__metadata": {"type": "SP.ListItem"},
            "FileLeafRef": new_folder_name,  # Nome visível no SharePoint
            "Title": new_folder_name        # Título da pasta
        }

        response = requests.post(url, headers=headers, data=json.dumps(data))

        if response.status_code == 200:
            print(f"Pasta {old_folder_name} renomeada com sucesso para {new_folder_name}.")
        else:
            print(f"Erro ao renomear a pasta {old_folder_name}: {response.status_code} - {response.text}")

    except Exception as e:
        print(f"Erro ao renomear a pasta {old_folder_name}: {e}")
        traceback.print_exc()

# Main atualizado para verificar e processar arquivos pendentes no início
if __name__ == "__main__":
    try:
        # Excluir e recriar a pasta oculta no início
        recriar_pasta_oculta()

        access_token = get_access_token()

        # Configurar tkinter para detectar o fechamento da janela
        root = tk.Tk()
        root.withdraw()  # Ocultar a janela principal temporariamente

    # Verificar se o termo já foi aceito
        if not verificar_aceitacao_termo():
            aceitou_termo = exibir_termo()
        if aceitou_termo:
            salvar_aceitacao_termo()
        else:
            messagebox.showinfo("Programa Encerrado", "Você precisa aceitar os termos para usar o programa.")
            sys.exit(0)

        root.deiconify()
        root.title("Captura de Prints")
        root.geometry("300x100")

        # Label informando que os prints estão sendo capturados
        label = tk.Label(root, text="Prints sendo disparados...")
        label.pack(pady=20)

        # Callback para fechamento da janela
        root.protocol("WM_DELETE_WINDOW", on_closing)

        current_date = datetime.now().strftime('%Y-%m-%d')

        # Pedir ao usuário para inserir o nome no Teams
        teams_name = simpledialog.askstring("Nome no Teams", "Digite o seu nome no Teams (deve terminar com 'AGIL LTDA'):", parent=root)
        teams_name = validar_nome_teams(teams_name)
        if not teams_name:
            messagebox.showerror("Erro", "O nome deve seguir o formato 'comercialX AGIL LTDA' (sem espaço entre 'comercial' e o número).")
            sys.exit(1)

        # Corrigir pastas com espaços no nome
        verificar_e_corrigir_pastas_incorretas(access_token)

        # Criar a pasta do usuário no SharePoint
        criar_ambiente_sharepoint(teams_name, current_date, access_token)

        # Definir o manipulador de sinal para detectar Ctrl+C e SIGTERM
        signal.signal(signal.SIGINT, signal_handler)
        signal.signal(signal.SIGTERM, signal_handler)

        # Iniciar captura de prints em um thread separado
        iniciar_captura_em_thread(access_token, teams_name, current_date)

        # Iniciar o loop da interface gráfica
        root.mainloop()

    except Exception as e:
        logging.error(f"Ocorreu um erro durante o processo principal: {e}")
        traceback.print_exc()
