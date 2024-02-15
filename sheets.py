import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = "1tXHai4DDZ2SdIEXnwrYwFMaot_MmIakQK01q0i31lV0"
SAMPLE_RANGE_NAME = "engenharia_de_software!A4:H"


def main():
  creds = None
  if os.path.exists("token.json"):
    creds = Credentials.from_authorized_user_file("token.json", SCOPES)
  if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
      creds.refresh(Request())
    else:
      flow = InstalledAppFlow.from_client_secrets_file(
          "credentials.json", SCOPES
      )
      creds = flow.run_local_server(port=0)

    with open("token.json", "w") as token:
      token.write(creds.to_json())

  try:
    service = build("sheets", "v4", credentials=creds)


    sheet = service.spreadsheets()
    result = (
        sheet.values()
        .get(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=SAMPLE_RANGE_NAME)
        .execute()
    )
    values = result.get("values", [])

    if not values:
      print("No data found.")
      return

    for row in values:
      falta = int(row[2])
      grade1 = float(row[3])
      grade2 = float(row[4])
      grade3 = float(row[5])
      media = round((grade1+grade2+grade3)/3)
      
      if falta >15:
        situacao = "Reprovado por falta"
        naf = 0
      elif media <50:
        situacao ="Reprovado por nota"
        naf = 0
      elif media >=70:
        situacao = "Aprovado"
        naf = 0
      elif media>=50 and media<70:
        naf = (100-media)
        situacao = "Exame"
      linha = int(row[0])+3
      update_values('G'+str(linha),'USER_ENTERED',situacao)
      update_values('H'+str(linha),'USER_ENTERED',naf)    

  except HttpError as err:
    print(err)





def update_values(range_name, value_input_option, _values):
  creds = Credentials.from_authorized_user_file("token.json", SCOPES)
  try:
    service = build("sheets", "v4", credentials=creds)
    values = [
        [
            _values
        ],

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
    print(f"{result.get('updatedCells')} cells updated.")
    return result
  except HttpError as error:
    print(f"An error occurred: {error}")
    return error




if __name__ == "__main__":
  main()