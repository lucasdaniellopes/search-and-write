from modules.api_utils import APIUtils

class SpreadsheetsUtils:
    @staticmethod
    def share_spreadsheet(spreadsheet, email, role='writer'):
        # Compartilha a planilha com o e-mail fornecido e define o papel (role)
        spreadsheet.share(email, perm_type='user', role=role)

    @staticmethod
    def find_spreadsheets_in_folder(drive_service, folder_id):
        sheets = []
        results = APIUtils.retry_api_call(
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
