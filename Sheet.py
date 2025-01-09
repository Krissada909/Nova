import gspread
from oauth2client.service_account import ServiceAccountCredentials

class GoogleSheetProcessor:
    
    def __init__(self, json_keyfile):
        self.json_keyfile = json_keyfile
        self.client = self.authenticate()
    
    def authenticate(self):
        """ Authenticate with Google Sheets API using service account. """
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_name(self.json_keyfile, scope)
        client = gspread.authorize(creds)
        return client
    
    def read_data_from_sheet(self, spreadsheet_id, sheet_name):
        """
        Reads data from a specific sheet in the Google Sheets file.
        
        :param spreadsheet_id: str, ID of the spreadsheet to access
        :param sheet_name: str, name of the sheet to read from
        :return: list, rows of data from the sheet
        """
        try:
            sheet = self.client.open_by_key(spreadsheet_id).worksheet(sheet_name)
            data = sheet.get_all_records()  # Retrieve all data in the sheet
            print(f"Read data from {sheet_name}: {data}")
            return data
        except Exception as e:
            print(f"Error reading data from sheet: {e}")
    
    def process_data(self, data):
        """
        Process the data read from the sheet (e.g., calculate averages, transform data, etc.)
        
        :param data: list, rows of data from the sheet
        :return: processed data
        """
        # Example processing: just return the data as is for now
        print(f"Processing data: {data}")
        return data  # Placeholder for actual data processing
    
    def write_data_to_sheet(self, spreadsheet_id, sheet_name, data):
        """
        Writes processed data back into the Google Sheets.
        
        :param spreadsheet_id: str, ID of the spreadsheet to update
        :param sheet_name: str, name of the sheet to write to
        :param data: list, data to write into the sheet
        """
        try:
            sheet = self.client.open_by_key(spreadsheet_id).worksheet(sheet_name)
            for i, row in enumerate(data, start=2):  # Start from row 2 (assuming header is in row 1)
                sheet.append_row(row)
            print(f"Successfully wrote data to {sheet_name}")
        except Exception as e:
            print(f"Error writing data to sheet: {e}")

# Example Usage
spreadsheet_id = "your_spreadsheet_id_here"
sheet_name = "Sheet1"
json_keyfile = "path_to_your_service_account_json_keyfile.json"

processor = GoogleSheetProcessor(json_keyfile)
data = processor.read_data_from_sheet(spreadsheet_id, sheet_name)
processed_data = processor.process_data(data)
processor.write_data_to_sheet(spreadsheet_id, sheet_name, processed_data)
