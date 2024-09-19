from flask import Flask, render_template, request, redirect, url_for
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import uuid
import os
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
import time

app = Flask(__name__)

# Configurações de autenticação
scope = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive.file",
    "https://www.googleapis.com/auth/drive"
]
creds = ServiceAccountCredentials.from_json_keyfile_name("C:/Users/Luize Zatti/Desktop/Metrificacao.json", scope)
client = gspread.authorize(creds)

# Abre a planilha
spreadsheet_url = "https://docs.google.com/spreadsheets/d/12NbcEHHlYBi9COhZfSo4kvvVNgXRD6YaMnSOEtLJS8I/edit"
spreadsheet = client.open_by_url(spreadsheet_url)
sheet = spreadsheet.sheet1

# Inicializa a API do Google Drive
drive_service = build('drive', 'v3', credentials=creds)

# ID da pasta onde os arquivos serão salvos
folder_id = '1u5PFDSfOTL5QHfnXgcLA-ErrJLJdRjAd'

@app.route('/')
def index():
    return render_template("index.html")

@app.route('/add_data', methods=['POST'])
def add_data():
    cliente = request.form['cliente']
    motivo = request.form['motivo']
    codigo_crm = request.form['codigo_crm']
    data_hora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    conversa = request.form['conversa']
    status = request.form['status']
    
    # Gera um ID único
    unique_id = str(uuid.uuid4())

    # Adiciona os dados à planilha, incluindo o ID único
    row = [unique_id, cliente, motivo, codigo_crm, data_hora, conversa, status]
    sheet.append_row(row)

    # Lida com o upload dos anexos
    if 'anexos' in request.files:
        files = request.files.getlist('anexos')
        for file in files:
            if file:
                # Salva o arquivo temporariamente
                file_path = os.path.join('temp', file.filename)
                file.save(file_path)

                # Faz o upload para o Google Drive
                media = MediaFileUpload(file_path, mimetype=file.content_type)
                drive_service.files().create(body={
                    'name': file.filename,
                    'parents': [folder_id]
                }, media_body=media, fields='id').execute()

                # Atraso antes de remover o arquivo
                time.sleep(1)

                # Remove o arquivo temporário
                try:
                    if os.path.exists(file_path):
                        os.remove(file_path)
                    else:
                        print(f"Arquivo {file_path} não encontrado.")
                except PermissionError:
                    print(f"Não foi possível remover {file_path} porque está sendo usado por outro processo.")

    return redirect(url_for('index'))

if __name__ == '__main__':
    if not os.path.exists('temp'):
        os.makedirs('temp')  # Cria a pasta temporária, se não existir
    app.run(debug=True)
