from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# The ID and range of a sample spreadsheet.
LULU_ORDER_SHEET = '1ovW5iLuGeadwMZa5_0ysAzzuZyLSaGYqI4ujYqZx7P8'
RANGE = 'A1:B19'

def main():
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('sheets', 'v4', credentials=creds)
    sheet = service.spreadsheets()

    # Set up dictionary to store counts for each
    shirtSizeOptions = ["Not Interested", "XXS", "XS", "S", "M", "L", "XL", "XXL"]
    sizeCountsDict = {}
    for option in shirtSizeOptions:
        sizeCountsDict[option] = 0

    # Get Data for column 1
    range1 = "D2:D19"
    result = sheet.values().get(spreadsheetId=LULU_ORDER_SHEET,
                                range=range1).execute()
    values = result.get('values', [])

    if not values:
        print('No data found.')
    else:
        # Store all sizes in array
        selectedSizes = []
        for row in values:
            selectedSizes.append(row[0])

    for selection in selectedSizes:
        if sizeCountsDict[selection] is not None:
            sizeCountsDict[selection] += 1

    print(sizeCountsDict)
    
     # Set up data to enter
    values = []
    clearValues = []
    body = {
        'values': values
    }
    clearBody = {
        'values': clearValues
    }

    i = 0
    for size, count in sizeCountsDict.items():
        values.append( ["{} - {}".format(size, count)] )
        clearValues.append( [""] )
        i += 1
    
    rangeStartNum = 28
    rangeEndNum = rangeStartNum + len(values)

    outputRange = 'D{}:D{}'.format(rangeStartNum, rangeEndNum)
    # Write Changes
    result = sheet.values().update(
    spreadsheetId=LULU_ORDER_SHEET, range=outputRange,
    valueInputOption="USER_ENTERED", body=body).execute()

    # Verify changes
    print('{0} cells updated.'.format(result.get('updatedCells')))

    # Get number of items (number of columns)
    range = "D2:AZ2"
    result = sheet.values().get(spreadsheetId=LULU_ORDER_SHEET,
                                range=range).execute()
    values = result.get('values', [])
    numRows = 0
    for row in values:
        print(row[0])
        if row[0] != "":
            
            numRows += 1
        else:
            break

    numItems = numRows - 2
    print("{} clothing items".format(numItems))

    DELETE = False
    if DELETE:
        rangeStartNum = 28
        rangeEndNum = rangeStartNum + len(values)

        outputRange = 'D{}:D{}'.format(rangeStartNum, rangeEndNum)
        # Write Changes
        result = sheet.values().update(
        spreadsheetId=LULU_ORDER_SHEET, range=outputRange,
        valueInputOption="USER_ENTERED", body=clearBody).execute()

        print('{0} cells cleared.'.format(result.get('updatedCells')))


if __name__ == '__main__':
    main()