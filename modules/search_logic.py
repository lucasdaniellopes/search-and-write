import time
from googleapiclient.discovery import build
from google.oauth2 import service_account
import gspread
import pandas as pd
import tkinter as tk

from modules.spreadsheet_utils import SpreadsheetsUtils

class SearchLogic(SpreadsheetsUtils):
    def __init__(self, gui_instance):
        super().__init__()
        self.gui_instance = gui_instance

    def start_search(self):
        print("Início do start_search")
        if not self.gui_instance:
            print("Gui instance não encontrada.")
            return

            try:
                # Limpa a área de resultados e atualiza o status
                self.gui_instance.result_text.delete(1.0, tk.END)
                self.gui_instance.retry_delay_seconds = 5
                self.gui_instance.requests_count = 0
                self.gui_instance.start_time = time.time()
                self.gui_instance.update_status_var("Lendo Dados")

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
                    self.gui_instance.result_text.insert(tk.END, 'Nenhuma pasta encontrada.\n')
                    self.gui_instance.update_status_var("Finalizado")
                else:
                    for folder in folders:
                        folder_name = folder["name"]
                        folder_id = folder["id"]
                        self.gui_instance.update_result_text(f'\nProcurando planilhas na pasta: {folder_name} ({folder_id})\n')
                        self.gui_instance.update_status_var("Esperando API")

                        sheets = SpreadsheetsUtils.find_spreadsheets_in_folder(drive_service, folder_id)

                        if not sheets:
                            # Nenhuma planilha encontrada na pasta
                            self.gui_instance.update_result_text('Nenhuma planilha encontrada na pasta.\n')
                            self.gui_instance.update_status_var("Finalizado")
                        else:
                            for sheet in sheets:
                                sheet_name = sheet["name"]
                                sheet_id = sheet["id"]
                                self.gui_instance.update_result_text(f'\nProcurando na planilha: {sheet_name} ({sheet_id})\n')
                                self.gui_instance.update_status_var("Esperando API")

                                try:
                                    spreadsheet = gc.open_by_key(sheet_id)

                                    for worksheet in spreadsheet.worksheets():
                                        if worksheet.title.upper().startswith(target_month):
                                            self.gui_instance.update_result_text(f'\nProcurando na aba: {worksheet.title}\n')
                                            self.gui_instance.update_status_var("Esperando API")

                                            cell_content = worksheet.get_all_values()
                                            df = pd.DataFrame(cell_content[1:], columns=cell_content[0])

                                            duplicate_columns = df.columns[df.columns.duplicated()].unique()
                                            df.columns = [f'{col}_{i}' if col in duplicate_columns else col for i, col in enumerate(df.columns)]

                                            df['Planilha Origem'] = sheet_name  # salvando nome da pasta

                                            filtered_df = df[df.apply(lambda row: any(search_value.upper() in str(cell).upper() for cell in row), axis=1)]

                                            if not filtered_df.empty:
                                                if new_spreadsheet is None:
                                                    # Cria uma nova planilha e a compartilha
                                                    self.gui_instance.update_result_text("Criando nova planilha...\n")
                                                    self.gui_instance.update_status_var("Esperando API")

                                                    new_spreadsheet = gc.create(f"Resultados_{search_value}")
                                                    SpreadsheetsUtils.share_spreadsheet(new_spreadsheet, 'relacionamento@hospitaldayunifip.com.br', role='writer')
                                                    self.gui_instance.update_result_text(f'Link da nova planilha: {new_spreadsheet.url}\n')
                                                    self.gui_instance.update_status_var("Escrevendo Dados")

                                                if not new_spreadsheet.get_worksheet(0).get_all_records():
                                                    # Adiciona uma linha vazia à nova planilha e a compartilha novamente
                                                    new_spreadsheet.get_worksheet(0).append_rows([['']] * 1)
                                                    SpreadsheetsUtils.share_spreadsheet(new_spreadsheet, 'relacionamento@hospitaldayunifip.com.br', role='writer')
                                                    self.gui_instance.update_result_text(f'Planilha compartilhada com sucesso.\n')
                                                    self.gui_instance.update_status_var("Esperando API")

                                                # Adiciona as linhas filtradas à nova planilha
                                                new_spreadsheet.get_worksheet(0).append_rows(filtered_df.values.tolist(),
                                                                                            value_input_option='USER_ENTERED',
                                                                                            table_range='A:Z')
                                                self.gui_instance.update_result_text("Linhas adicionadas à nova planilha.\n")
                                                self.gui_instance.update_status_var("Esperando API")

                                except Exception as e:
                                    self.gui_instance.update_result_text(f'Ocorreu um erro: {e}\n')
                                    self.gui_instance.update_status_var("Finalizado com Erro")

            except Exception as e:
                self.gui_instance.update_result_text(f'Ocorreu um erro: {e}\n')
                self.gui_instance.update_status_var("Finalizado com Erro")
