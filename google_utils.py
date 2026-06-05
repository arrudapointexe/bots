import gspread
from oauth2client.service_account import ServiceAccountCredentials

def get_gspread_client():
    """
    Autoriza e retorna um cliente gspread para interagir com o Google Sheets.
    """
    escopo = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    credenciais = ServiceAccountCredentials.from_json_keyfile_name('credenciais.json', escopo)
    cliente = gspread.authorize(credenciais)
    return cliente
