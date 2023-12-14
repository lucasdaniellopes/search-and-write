import time
from googleapiclient.discovery import build
from google.oauth2 import service_account
import gspread
import pandas as pd
import tkinter as tk
from tkinter import ttk, scrolledtext

def retry_api_call(api_function, *args, **kwargs):
    max_retries = 3
    retry_delay_seconds = 5

    quota_exceeded_count = 0  # Contador de limites de cota excedidos

    for retry_count in range(max_retries):
        try:
            result = api_function(*args, **kwargs)
            return result
        except Exception as e:
            if 'RATE_LIMIT_EXCEEDED' in str(e):
                print(f'Taxa limite excedida. Tentando novamente em {retry_delay_seconds} segundos...')
                time.sleep(retry_delay_seconds)
            elif '429' in str(e):  # Verifica se o código de erro é 429 (limite de cota excedido)
                quota_exceeded_count += 1
                print(f'Limite de cota excedido. Contador: {quota_exceeded_count}')
                time.sleep(retry_delay_seconds)
            else:
                raise

    raise Exception(f'Não foi possível realizar a chamada após {max_retries} tentativas.')

def create_gui():
    root = tk.Tk()
    root.title("Pesquisa de Planilhas")

    search_label = ttk.Label(root, text="Digite o valor a ser procurado nas planilhas:")
    search_label.pack(pady=5)

    search_entry = ttk.Entry(root)
    search_entry.pack(pady=5)

    month_label = ttk.Label(root, text="Digite o mês a ser procurado (por exemplo, DEZEMBRO):")
    month_label.pack(pady=5)

    month_entry = ttk.Entry(root)
    month_entry.pack(pady=5)

    result_text = scrolledtext.ScrolledText(root, width=40, height=10)
    result_text.pack(pady=10)

    quota_count_label = ttk.Label(root, text="Contador de Limite de Cota Excedido:")
    quota_count_label.pack(pady=5)

    quota_count_var = tk.StringVar()
    quota_count_var.set("0")
    quota_count_display = ttk.Label(root, textvariable=quota_count_var)
    quota_count_display.pack(pady=5)

    def start_search():
        nonlocal result_text, quota_count_var
        search_value = search_entry.get()
        target_month = month_entry.get().upper()

        try:
            result_text.delete(1.0, tk.END)

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

            new_spreadsheet = None

            folders_query = "mimeType='application/vnd.google-apps.folder' and trashed=false"
            folders_results = drive_service.files().list(q=folders_query, fields="files(id, name)").execute()
            folders = folders_results.get('files', [])

            if not folders:
                result_text.insert(tk.END, 'Nenhuma pasta encontrada.\n')
                root.update()  # Atualiza a interface gráfica
            else:
                for folder in folders:
                    folder_name = folder["name"]
                    folder_id = folder["id"]
                    result_text.insert(tk.END, f'\nProcurando planilhas na pasta: {folder_name} ({folder_id})\n')
                    root.update()  # Atualiza a interface gráfica
                    result_text.see(tk.END)  # Rola para a última linha

                    sheets = find_spreadsheets_in_folder(folder_id)

                    if not sheets:
                        result_text.insert(tk.END, 'Nenhuma planilha encontrada na pasta.\n')
                        root.update()  # Atualiza a interface gráfica
                    else:
                        for sheet in sheets:
                            sheet_name = sheet["name"]
                            sheet_id = sheet["id"]
                            result_text.insert(tk.END, f'\nProcurando na planilha: {sheet_name} ({sheet_id})\n')
                            root.update()  # Atualiza a interface gráfica
                            result_text.see(tk.END)  # Rola para a última linha

                            try:
                                spreadsheet = gc.open_by_key(sheet_id)

                                for worksheet in spreadsheet.worksheets():
                                    if worksheet.title.upper().startswith(target_month):
                                        result_text.insert(tk.END, f'\nProcurando na aba: {worksheet.title}\n')
                                        root.update()  # Atualiza a interface gráfica
                                        result_text.see(tk.END)  # Rola para a última linha

                                        cell_content = worksheet.get_all_values()
                                        df = pd.DataFrame(cell_content[1:], columns=cell_content[0])

                                        # Tornar os nomes das colunas únicos
                                        duplicate_columns = df.columns[df.columns.duplicated()].unique()
                                        df.columns = [f'{col}_{i}' if col in duplicate_columns else col for i, col in enumerate(df.columns)]

                                        df['Planilha Origem'] = sheet_name

                                        filtered_df = df[df.apply(lambda row: any(search_value.upper() in str(cell).upper() for cell in row), axis=1)]

                                        if not filtered_df.empty:
                                            if new_spreadsheet is None:
                                                result_text.insert(tk.END, "Criando nova planilha...\n")
                                                root.update()  # Atualiza a interface gráfica
                                                result_text.see(tk.END)  # Rola para a última linha

                                                new_spreadsheet = gc.create(f"Resultados_{search_value}")
                                                share_spreadsheet(new_spreadsheet, 'relacionamento@hospitaldayunifip.com.br', role='writer')
                                                result_text.insert(tk.END, f'Link da nova planilha: {new_spreadsheet.url}\n')
                                                root.update()  # Atualiza a interface gráfica
                                                result_text.see(tk.END)  # Rola para a última linha

                                            if not new_spreadsheet.get_worksheet(0).get_all_records():
                                                new_spreadsheet.get_worksheet(0).append_rows([['']] * 1)

                                            new_spreadsheet.get_worksheet(0).append_rows(filtered_df.values.tolist(), value_input_option='USER_ENTERED', table_range='A:Z')
                                            result_text.insert(tk.END, "Linhas adicionadas à nova planilha.\n")
                                            root.update()  # Atualiza a interface gráfica
                                            result_text.see(tk.END)  # Rola para a última linha

                            except Exception as e:
                                if '429' in str(e):
                                    quota_exceeded_count += 1
                                    quota_count_var.set(str(quota_exceeded_count))  # Atualiza o contador na interface gráfica
                                result_text.insert(tk.END, f'Ocorreu um erro: {e}\n')
                                root.update()  # Atualiza a interface gráfica
                                result_text.see(tk.END)  # Rola para a última linha

                            time.sleep(2.5)

        except Exception as e:
            result_text.insert(tk.END, f'Ocorreu um erro: {e}\n')
            result_text.see(tk.END)  # Rola para a última linha

    search_button = ttk.Button(root, text="Iniciar Pesquisa", command=start_search)
    search_button.pack(pady=10)

    root.mainloop()

# Criar a interface gráfica
create_gui()
