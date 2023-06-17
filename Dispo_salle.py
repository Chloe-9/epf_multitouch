import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font
from datetime import datetime
from msal import ConfidentialClientApplication
import requests

# Paramètres d'authentification
client_id = 'ad6f3bec-ace3-47fd-8abd-e09681001eb2'
client_secret = '1vB8Q~iyulieE843cnrWBYxdvS_l.7dqc2oTNdzH'
tenant_id = '07c72520-e645-4679-8f55-dc611c975a02'

# Configuration de l'application Microsoft Graph
authority = f'https://login.microsoftonline.com/{tenant_id}'
scope = ['https://graph.microsoft.com/.default']

# Connexion et récupération du token d'accès
app = ConfidentialClientApplication(client_id, authority=authority, client_credential=client_secret)
token = app.acquire_token_silent(scope, account=None)['access_token']

# Requête pour récupérer les événements du calendrier
url = 'https://graph.microsoft.com/v1.0/me/events'
headers = {'Authorization': 'Bearer ' + token}
params = {'$select': 'subject,start,end,location'}

response = requests.get(url, headers=headers, params=params)
events = response.json().get('value', [])

# Création d'un nouveau classeur Excel
wb = openpyxl.Workbook()
ws = wb.active

# En-têtes de colonnes
headers = ['Subject', 'Start', 'End', 'Location']
for col_num, header in enumerate(headers, 1):
    col_letter = get_column_letter(col_num)
    ws[f'{col_letter}1'] = header
    ws[f'{col_letter}1'].font = Font(bold=True)

# Remplissage des données des événements
for row_num, event in enumerate(events, 2):
    ws[f'A{row_num}'] = event['subject']
    ws[f'B{row_num}'] = event['start'].get('dateTime')
    ws[f'C{row_num}'] = event['end'].get('dateTime')
    ws[f'D{row_num}'] = event['location'].get('displayName')

# Enregistrement du fichier Excel
file_name = f'calendrier_{datetime.now().strftime("%Y-%m-%d_%H-%M-%S")}.xlsx'
wb.save(file_name)
print(f"Les événements ont été exportés dans le fichier : {file_name}")

