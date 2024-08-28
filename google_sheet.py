import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import json

class GoogleSheet:

    def __init__(self, docid, worksheet_name, credentials_string):
        credentials = json.loads(credentials_string, strict=False)
        client = gspread.service_account_from_dict(credentials)
        sh = client.open_by_key(docid)
        self.worksheet = sh.worksheet(worksheet_name)

    def setGoogleSheet(self, data):
        print("Uploading Google Sheet.")
        # Convert DataFrame to list of lists
        values = data.values.tolist()
        
        # Add headers to the values
        values.insert(0, data.columns.tolist())
        
        # Define the range to update
        num_rows = len(values)
        num_columns = len(values[0])
        cells_range = f"A1:{chr(65 + num_columns - 1)}{num_rows}"
        
        # Update cells in the sheet
        self.worksheet.update(range_name=cells_range, values=values)

    def blankGoogleSheet(self):
        print ("Blanking Google Sheet.")
        sheet = self.worksheet.get_all_values()
        headers = sheet.pop(0)
        sheet_df = pd.DataFrame(sheet, columns=headers)
        
        # Number of rows and columns
        num_rows = len(sheet_df.index) + 1  # +1 for the header row
        num_columns = len(sheet_df.columns)

        if num_columns>0:        
            # Define the range of cells to blank
            cells = f"A1:{chr(65 + num_columns - 1)}{num_rows}"
            cell_list = self.worksheet.range(cells)
            
            # Blank out the cells
            for cell in cell_list:
                cell.value = ""
            
            # Update the cells in the sheet
            self.worksheet.update_cells(cell_list)
        
        else:
            print ("Worksheet was already blank.")

    def getGoogleSheet(self):
        print ("Downloading Google Sheet.")
        #Get Form Responses tab as a list of lists
        sheet = self.worksheet.get_all_values()

        #Convert sheet to dataframe
        headers = sheet.pop(0)
        sheet = pd.DataFrame(sheet, columns=headers)

        return sheet

    def getNewRows(self, new, old):
        print ("Getting new rows.")
        return new[~new.index.isin(old.index)]
    
    def appendRow(self, row):
        self.worksheet.append_row(row)

json_credential = """
{
    "type": "service_account",
    "project_id": "",
    "private_key_id": "",
    "private_key": "",
    "client_email": "",
    "client_id": "",
    "auth_uri": "",
    "token_uri": "",
    "auth_provider_x509_cert_url": "",
    "client_x509_cert_url": ""
}"""

example = pd.read_csv("example.csv")

#Sample URL: https://docs.google.com/spreadsheets/d/1NbE4VXFA1fhkwql2DjzeEwfgrfW7B3Uj4omyfOMwOYE/edit
#docid is the string of numbers after the "/d/" and before the "/edit"
docid = "1NbE4VXFA1dgxwql2DjzeEwmkrfW7B3Uj4omyfOMwOYE"

#worksheet_name is the name of the tab
worksheet_name = "Class Data"

sheet = GoogleSheet(docid, worksheet_name, json_credential)

#Blank the Google Sheet
sheet.blankGoogleSheet()

#Write new dataframe to the Google Sheet
sheet.setGoogleSheet(example)

#Get the contents of the Google Sheet as a DataFrame
df = sheet.getGoogleSheet()
print (df)