from googleapiclient.discovery import build
from google.oauth2 import service_account
import gspread
from gspread_dataframe import get_as_dataframe
import pandas as pd

# Carregue as credenciais da conta de serviço
credentials = service_account.Credentials.from_service_account_file(
    'chave-da-conta-de-servico.json',
    scopes=['https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/spreadsheets']
)

# Crie uma instância do serviço Google Drive
drive_service = build('drive', 'v3', credentials=credentials)

# Conecte-se ao Google Sheets usando gspread
gc = gspread.service_account(filename='chave-da-conta-de-servico.json')

def find_spreadsheets_in_folder(folder_id):
    sheets = []

    # Obter todos os arquivos na pasta
    results = drive_service.files().list(
        q=f"'{folder_id}' in parents and trashed=false",
        fields="files(id, name, mimeType)"
    ).execute()

    files = results.get('files', [])

    for file in files:
        file_id = file['id']
        file_name = file['name']
        mime_type = file.get('mimeType', '')

        # Verificar se o arquivo é uma planilha
        if 'spreadsheet' in mime_type:
            sheets.append({'id': file_id, 'name': file_name})

    return sheets

try:
    # Liste as pastas na raiz do Google Drive
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

            # Obter planilhas na pasta usando a função definida acima
            sheets = find_spreadsheets_in_folder(folder_id)

            if not sheets:
                print('Nenhuma planilha encontrada na pasta.')
            else:
                # Agora, para cada planilha, faça o processamento
                for sheet in sheets:
                    sheet_name = sheet["name"]
                    sheet_id = sheet["id"]
                    print(f'\nProcurando na planilha: {sheet_name} ({sheet_id})')

                    try:
                        # Leitura dos dados da planilha usando gspread
                        spreadsheet = gc.open_by_key(sheet_id)
                        worksheet = spreadsheet.sheet1  # Assumindo que estamos usando a primeira folha
                        cell_content = worksheet.get_all_values()

                        # Converte o conteúdo da célula para DataFrame
                        df = pd.DataFrame(cell_content[1:], columns=cell_content[0])

                        # Filtrar linhas com "XRAY (2)"
                        filtered_df = df[df.apply(lambda row: "XRAY (2)" in str(row).upper(), axis=1)]

                        if not filtered_df.empty:
                            print(f'Dados encontrados na planilha "{sheet_name}":')
                            print(filtered_df)

                    except gspread.exceptions.APIError as e:
                        # Verificar se o erro é 'FAILED_PRECONDITION'
                        if 'FAILED_PRECONDITION' in str(e):
                            print(f'Ignorando erro de pré-condição falhada para a planilha "{sheet_name}"')
                            continue
                        else:
                            raise  # Propagar outros erros

except Exception as e:
    print(f'Ocorreu um erro: {e}')