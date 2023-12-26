import time
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.oauth2 import service_account
import gspread
import pandas as pd
import tkinter as tk
from tkinter import ttk, scrolledtext


class PlanilhaSearchApp:
    def __init__(self):
        # Configuração da interface gráfica
        self.root = tk.Tk()
        self.root.title("Pesquisa de Planilhas")

        self.search_label = ttk.Label(self.root, text="Digite o valor a ser procurado nas planilhas:")
        self.search_label.pack(pady=5)

        self.search_entry = ttk.Entry(self.root)
        self.search_entry.pack(pady=5)

        self.month_label = ttk.Label(self.root, text="Digite o mês a ser procurado (por exemplo, DEZEMBRO):")
        self.month_label.pack(pady=5)

        self.month_entry = ttk.Entry(self.root)
        self.month_entry.pack(pady=5)

        self.result_text = scrolledtext.ScrolledText(self.root, width=40, height=10)
        self.result_text.pack(pady=10)

        self.status_label = ttk.Label(self.root, text="Status da Aplicação:")
        self.status_label.pack(pady=5)

        self.status_var = tk.StringVar()
        self.status_var.set("Aguardando Início")
        self.status_display = ttk.Label(self.root, textvariable=self.status_var)
        self.status_display.pack(pady=5)

        self.retry_delay_seconds = 5
        self.max_wait_time_seconds = 180
        self.requests_count = 0
        self.start_time = time.time()

        # Inicia a interface gráfica
        self.create_gui()

    def retry_api_call(self, api_function, *args, **kwargs):
        last_request = None  # Reinicializa a variável
        retry_delay_seconds = 5  # Inicializa o valor de espera

        while self.requests_count < 60:
            elapsed_time = time.time() - self.start_time

            if elapsed_time >= 60:
                self.requests_count = 0
                self.start_time = time.time()

            try:
                result = api_function(*args, **kwargs)
                self.requests_count += 1
                print("QUASE LAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA")
                return result
            except Exception as e:
                if 'RATE_LIMIT_EXCEEDED' in str(e) or (hasattr(e, 'code') and e.code == 429):
                    print("ENTROOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOU")
                    print(f'Erro: {str(e)}. Aguardando...')

                    # Aplica espera exponencial antes de reenviar a solicitação
                    time.sleep(retry_delay_seconds)

                    # Aumenta o tempo de espera exponencial para a próxima tentativa
                    retry_delay_seconds *= 2
                    continue
                else:
                    print(f'Erro: {str(e)}. Tentando novamente...')
                    last_request = {'api_function': api_function, 'args': args, 'kwargs': kwargs}
                    retry_delay_seconds *= 2
                    continue

        # Se chegou a este ponto, todas as tentativas foram esgotadas
        # Tentar novamente a última chamada que falhou
        if last_request:
            print('Tentando novamente a última chamada que falhou...')
            self.retry_api_call(last_request['api_function'], *last_request['args'], **last_request['kwargs'])
        else:
            print("Limite de tentativas atingido. Não foi possível realizar a chamada.")

    def share_spreadsheet(self, spreadsheet, email, role='writer'):
        # Compartilha a planilha com o e-mail fornecido e define o papel (role)
        spreadsheet.share(email, perm_type='user', role=role)

    def find_spreadsheets_in_folder(self, drive_service, folder_id):
        sheets = []
        results = self.retry_api_call(
            drive_service.files().list,
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

    def update_status_var(self, message):
        # Atualiza a variável de status e a interface gráfica
        self.status_var.set(message)
        self.root.update()

    def update_result_text(self, message):
        # Atualiza a área de resultados e a interface gráfica
        self.result_text.insert(tk.END, message)
        self.root.update()

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

    def create_gui(self):
        # Criação da interface gráfica
        search_button = ttk.Button(self.root, text="Iniciar Pesquisa", command=self.start_search)
        search_button.pack(pady=10)

        stop_button = ttk.Button(self.root, text="Encerrar", command=self.root.destroy)
        stop_button.pack(pady=10)

        # Inicia a interface gráfica
        self.root.mainloop()


if __name__ == "__main__":
    # Executa a aplicação
    app = PlanilhaSearchApp()
