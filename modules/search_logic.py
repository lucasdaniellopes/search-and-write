import time
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.oauth2 import service_account
import gspread
import pandas as pd
import tkinter as tk


from modules.api_utils import APIUtils
from modules.spreadsheet_utils import SpreadsheetUtils
from modules.gui import GUI

class SearchLogic(APIUtils, SpreadsheetUtils, GUI):
    def __init__(self):
        super().__init__()

    @staticmethod
    def start_search(self):
        # Função principal para iniciar a pesquisa
        search_value = self.search_entry.get()
        target_month = self.month_entry.get().upper()

        try:
            # Limpa a área de resultados e atualiza o status
            self.result_text.delete(1.0, tk.END)
            self.retry_delay_seconds = 5
            self.requests_count = 0
            self.start_time = time.time()
            self.update_status_var("Lendo Dados")

            # Configuração das credenciais para acessar o Google Drive e Google Sheets
            credentials = service_account.Credentials.from_service_account_file(
                'chave-da-conta-de-servico.json',
                scopes=['https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/spreadsheets']
            )

            drive_service = build('drive', 'v3', credentials=credentials)
            gc = gspread.service_account(filename='chave-da-conta-de-servico.json')

            new_spreadsheet = None

            # Query para encontrar pastas no Google Drive
            folders_query = "mimeType='application/vnd.google-apps.folder' and trashed=false"
            folders_results = drive_service.files().list(q=folders_query, fields="files(id, name)").execute()
            folders = folders_results.get('files', [])

            if not folders:
                # Nenhuma pasta encontrada
                self.result_text.insert(tk.END, 'Nenhuma pasta encontrada.\n')
                self.update_status_var("Finalizado")
            else:
                for folder in folders:
                    folder_name = folder["name"]
                    folder_id = folder["id"]
                    self.update_result_text(f'\nProcurando planilhas na pasta: {folder_name} ({folder_id})\n')
                    self.update_status_var("Esperando API")

                    sheets = self.find_spreadsheets_in_folder(drive_service, folder_id)

                    if not sheets:
                        # Nenhuma planilha encontrada na pasta
                        self.update_result_text('Nenhuma planilha encontrada na pasta.\n')
                        self.update_status_var("Finalizado")
                    else:
                        for sheet in sheets:
                            sheet_name = sheet["name"]
                            sheet_id = sheet["id"]
                            self.update_result_text(f'\nProcurando na planilha: {sheet_name} ({sheet_id})\n')
                            self.update_status_var("Esperando API")

                            try:
                                spreadsheet = gc.open_by_key(sheet_id)

                                for worksheet in spreadsheet.worksheets():
                                    if worksheet.title.upper().startswith(target_month):
                                        self.update_result_text(f'\nProcurando na aba: {worksheet.title}\n')
                                        self.update_status_var("Esperando API")

                                        cell_content = worksheet.get_all_values()
                                        df = pd.DataFrame(cell_content[1:], columns=cell_content[0])

                                        duplicate_columns = df.columns[df.columns.duplicated()].unique()
                                        df.columns = [f'{col}_{i}' if col in duplicate_columns else col for i, col in enumerate(df.columns)]

                                        df['Planilha Origem'] = sheet_name #salvando nome da pasta

                                        filtered_df = df[df.apply(lambda row: any(search_value.upper() in str(cell).upper() for cell in row), axis=1)]

                                        if not filtered_df.empty:
                                            if new_spreadsheet is None:
                                                # Cria uma nova planilha e a compartilha
                                                self.update_result_text("Criando nova planilha...\n")
                                                self.update_status_var("Esperando API")

                                                new_spreadsheet = gc.create(f"Resultados_{search_value}")
                                                self.share_spreadsheet(new_spreadsheet, 'relacionamento@hospitaldayunifip.com.br', role='writer')
                                                self.update_result_text(f'Link da nova planilha: {new_spreadsheet.url}\n')
                                                self.update_status_var("Escrevendo Dados")

                                            if not new_spreadsheet.get_worksheet(0).get_all_records():
                                                # Adiciona uma linha vazia à nova planilha e a compartilha novamente
                                                new_spreadsheet.get_worksheet(0).append_rows([['']] * 1)
                                                self.share_spreadsheet(new_spreadsheet, 'relacionamento@hospitaldayunifip.com.br', role='writer')
                                                self.update_result_text(f'Planilha compartilhada com sucesso.\n')
                                                self.update_status_var("Esperando API")

                                            # Adiciona as linhas filtradas à nova planilha
                                            new_spreadsheet.get_worksheet(0).append_rows(filtered_df.values.tolist(), value_input_option='USER_ENTERED', table_range='A:Z')
                                            self.update_result_text("Linhas adicionadas à nova planilha.\n")
                                            self.update_status_var("Esperando API")

                            except Exception as e:
                                self.update_result_text(f'Ocorreu um erro: {e}\n')
                                self.update_status_var("Finalizado com Erro")

        except Exception as e:
            self.update_result_text(f'Ocorreu um erro: {e}\n')
            self.update_status_var("Finalizado com Erro")