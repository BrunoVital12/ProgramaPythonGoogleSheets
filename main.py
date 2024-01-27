import os.path
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

def main():
    creds = None

    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)

    # If there are no (valid) credentials available, let the user log in (the credentials are unique and not shareable)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
    # You need to rename your credentials file to client_secret.json      
            flow = InstalledAppFlow.from_client_secrets_file("client_secret.json", SCOPES)
            creds = flow.run_local_server(port=0)

        # Save the credentials for the next run
        with open("token.json", "w") as token:
            token.write(creds.to_json())

    try:
        service = build("sheets", "v4", credentials=creds)

        # Call the Sheets API
        sheet = service.spreadsheets()
        result = sheet.values().get(
            spreadsheetId="16_n92XWvhp7VSjwwMLMS4QxwmPA_k42Kf5Mnel6VDWs",
            range="engenharia_de_software!A3:F27"
        ).execute()

        values = result.get("values", [])
        for row in values:
            print(row)

        # Store the values ​​to be added
        valores_adicionar = [["Situação"]]
        valores_adicionar2 = [["Nota para Aprovação Final"]]

        # Loop through each row in values
        for i, row in enumerate(values):
            # Skip the header row (i > 0)
            if i > 0:
                # Setting the fault columns
                faltas = int(row[2])

                # Setting the grades columns
                p1 = int(row[3])
                p2 = int(row[4])
                p3 = int(row[5])

                # Doing the calculations
                media = (p1 / 10) + (p2 / 10) + (p3 / 10)
                finalmedia = media / 3

                # If structure to set approved or disapproved by presence
                if faltas > 15:
                    situacaofalta = 'Reprovado por falta'
                    naf = 0
                else:
                    situacaofalta = 'Aprovado'

                # If structure. 0 - approved. 1 - reproved by grades. 2 - final exam
                if situacaofalta == 'Aprovado' and finalmedia >= 7:
                    situacaofinal = 0
                elif situacaofalta == 'Aprovado' and finalmedia < 5:
                    situacaofinal = 1
                elif situacaofalta == 'Aprovado' and 5 <= finalmedia < 7:
                    situacaofinal = 2
                else:
                    situacaofinal = 'Situação final não prevista'

                # Setting the conditions approved or disapproved
                if situacaofinal == 0:
                    fim = 'Aprovado'
                    naf = 0
                    print(fim)
                elif situacaofinal == 1:
                    fim = 'Reprovado por Nota'
                    naf = 0
                    print(fim)
                elif situacaofalta == 'Reprovado por falta':
                    fim = 'Reprovado por Falta'
                    naf = 0
                    print(fim)
                elif situacaofinal == 2:
                    
                    # Calculate the final exam grade
                    naf = (5 * 2) - finalmedia
                    fim = 'Exame final'
                    print("{}. Nota para Aprovação final: {:.1f}".format(fim, naf))

                # Adding values to the list
                valores_adicionar.append([fim])
                valores_adicionar2.append([round(naf, 1)]) 

        # Updating the spreadsheet
        result = sheet.values().update(
            spreadsheetId="16_n92XWvhp7VSjwwMLMS4QxwmPA_k42Kf5Mnel6VDWs",
            range="engenharia_de_software!G3",
            valueInputOption="USER_ENTERED",
            body={"values": valores_adicionar}
        ).execute()

        result = sheet.values().update(
            spreadsheetId="16_n92XWvhp7VSjwwMLMS4QxwmPA_k42Kf5Mnel6VDWs",
            range="engenharia_de_software!H3",
            valueInputOption="USER_ENTERED",
            body={"values": valores_adicionar2}
        ).execute()

    except HttpError as err:
        print(err)


if __name__ == "__main__":
    main()
