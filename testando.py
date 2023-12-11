import time
import json
import math
from googleapiclient.discovery import build
from google.oauth2 import service_account
import gspread
import pandas as pd

def call_api_with_retry(api_function, *args, **kwargs):
    max_retries = 3
    retry_delay_seconds = 5

    for retry_count in range(max_retries):
        try:
            result = api_function(*args, **kwargs)
            return result
        except Exception as e:
            if 'RATE_LIMIT_EXCEEDED' in str(e):
                print(f'Taxa limite excedida. Tentando novamente em {retry_delay_seconds} segundos...')
                time.sleep(retry_delay_seconds)
            else:
                raise

    raise Exception(f'Não foi possível realizar a chamada após {max_retries} tentativas.')

credentials = service_account.Credentials.from_service_account_file(
    'chave-da-conta-de-servico.json',
    scopes=['https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/spreadsheets']
)

drive_service = build('drive', 'v3', credentials=credentials)
gc = gspread.service_account(filename='chave-da-conta-de-servico.json')

def share_spreadsheet(spreadsheet, email, role='writer'):
    spreadsheet.share(email, perm_type='user', role=role)

def find_spreadsheets_in_folder(folder_id):
    sheets = []
    results = drive_service.files().list(
        q=f"'{folder_id}' in parents and trashed=false",
        fields="files(id, name, mimeType)"
    ).execute()

    files = results.get('files', [])

    for file in files:
        file_id = file['id']
        file_name = file['name']
        mime_type = file.get('mimeType', '')
        if 'spreadsheet' in mime_type:
            sheets.append({'id': file_id, 'name': file_name})

    return sheets

# Variável para armazenar a nova planilha
new_spreadsheet = None

try:
    search_value = input("Digite o valor a ser procurado nas planilhas: ")
    target_month = input("Digite o mês a ser procurado (por exemplo, DEZEMBRO): ").upper()

    folders_query = "mimeType='application/vnd.google-apps.folder' and trashed=false"
    folders_results = drive_service.files().list(q=folders_query, fields="files(id, name)").execute()
    folders = folders_results.get('files', [])

    if not folders:
        print('Nenhuma pasta encontrada.')
    else:
        for folder in folders:
            folder_name = folder["name"]
            folder_id = folder["id"]
            print(f'\nProcurando planilhas na pasta: {folder_name} ({folder_id})')

            sheets = find_spreadsheets_in_folder(folder_id)

            if not sheets:
                print('Nenhuma planilha encontrada na pasta.')
            else:
                for sheet in sheets:
                    sheet_name = sheet["name"]
                    sheet_id = sheet["id"]
                    print(f'\nProcurando na planilha: {sheet_name} ({sheet_id})')

                    try:
                        spreadsheet = gc.open_by_key(sheet_id)

                        for worksheet in spreadsheet.worksheets():
                            if worksheet.title.upper().startswith(target_month):
                                print(f'\nProcurando na aba: {worksheet.title}')

                                cell_content = worksheet.get_all_values()
                                df = pd.DataFrame(cell_content[1:], columns=cell_content[0])

                                # Filtra linhas com o valor fornecido pelo usuário
                                filtered_df = df[df.apply(lambda row: any(search_value.upper() in str(cell).upper() for cell in row), axis=1)]

                                if not filtered_df.empty:
                                    if new_spreadsheet is None:
                                        print("Criando nova planilha...")
                                        new_spreadsheet = gc.create(f"Resultados_{search_value}")
                                        share_spreadsheet(new_spreadsheet, 'relacionamento@hospitaldayunifip.com.br', role='writer')
                                        print(f'Link da nova planilha: {new_spreadsheet.url}')

                                    # Adiciona o título da planilha pai como uma nova coluna
                                    filtered_df['Planilha Pai'] = sheet_name

                                    # Atualizando a nova planilha com os dados filtrados
                                    new_spreadsheet.get_worksheet(0).update([filtered_df.columns.values.tolist()] + filtered_df.values.tolist())
                                    print("Linhas adicionadas à nova planilha.")
                                    time.sleep(1)

                    except Exception as e:
                        print(f'Ocorreu um erro: {e}')

except Exception as e:
    print(f'Ocorreu um erro: {e}')
