import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

#Permissions within the API
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# The ID and range of the spreadsheet
SAMPLE_SPREADSHEET_ID = "1tXHai4DDZ2SdIEXnwrYwFMaot_MmIakQK01q0i31lV0"
SAMPLE_RANGE_NAME = "engenharia_de_software!A4:H"

# Checking credentials and validating credentials
def main():
  creds = None
  #Check if file token.json exists
  if os.path.exists("token.json"):
    creds = Credentials.from_authorized_user_file("token.json", SCOPES)
  if not creds or not creds.valid:
#Check the credentials.json file to valid acess
    if creds and creds.expired and creds.refresh_token:
      creds.refresh(Request())
    else:
      flow = InstalledAppFlow.from_client_secrets_file(
          "credentials.json", SCOPES
      )
      creds = flow.run_local_server(port=0)
# Creates file token.json with the file credentials.json to access the spreadsheet
    with open("token.json", "w") as token:
      token.write(creds.to_json())
#Read the spreadsheets
  try:
    service = build("sheets", "v4", credentials=creds)
    sheet = service.spreadsheets()
    result = (
        sheet.values()
        .get(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=SAMPLE_RANGE_NAME)
        .execute()
    )
    values = result.get("values", [])
#Check if spreadsheet is empty
    if not values:
      print("No data found.")
      return
#Caculations for Faltas and Notas
    for row in values:
      line = str(row[0])
      name = str(row[1])
      falta = int(row[2])
      grade1 = float(row[3])
      grade2 = float(row[4])
      grade3 = float(row[5])
      media = round((grade1+grade2+grade3)/3)
      #Check the status of Sitação and Nota para Aprovação Final
      if falta >15:
        situacao = "Reprovado por falta"
        naf = 0
      elif media <50:
        situacao ="Reprovado por nota"
        naf = 0
      elif media >=70:
        situacao = "Aprovado por nota"
        naf = 0
      elif media>=50 and media<70:
        naf = (100-media)
        situacao = "Exame final"
      linha = int(row[0])+3
      #Update the spreadsheet with the new cell values for Situação and Nota para Aprovação Final
      update_values('G'+str(linha),'USER_ENTERED',situacao)
      update_values('H'+str(linha),'USER_ENTERED',naf)
      print(line + " - " + name +" - "+situacao+" - "+str(naf))
  except HttpError as err:
    print(err)

#Update cells function
def update_values(range_name, value_input_option, _values):
  #Check credentials
  creds = Credentials.from_authorized_user_file("token.json", SCOPES)
  try:
    service = build("sheets", "v4", credentials=creds)
    values = [
        [
            _values
        ],
#Define input spreadsheet,range and type
    ]
    body = {"values": values}
    result = (
        service.spreadsheets()
        .values()
        .update(
            spreadsheetId=SAMPLE_SPREADSHEET_ID,
            range=range_name,
            valueInputOption=value_input_option,
            body=body,
          )
        .execute()
    )

    #Print cells values updated with name, Situação + Nota para Aprovação Final
    return result
  except HttpError as error:
    print(f"An error occurred: {error}")
    return error



#Runs code
if __name__ == "__main__":
  main()