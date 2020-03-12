from __future__ import print_function
import pickle
import os.path
import sys, getopt
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

DELETE = True

def main():
    # Parse args
    if len(sys.argv) < 2:
        print("USAGE: python lulu.py googleSheetURL")
        exit(1)
    url = sys.argv[1]
    LULU_ORDER_SHEET = url.split("/")[5]
    
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

    ## Verify changes
    #print('{0} cells updated.'.format(result.get('updatedCells')))

    # Get number of columns
    range1 = 'D2:AZ2'
    result = sheet.values().get(spreadsheetId=LULU_ORDER_SHEET,
                                range=range1).execute()
    values = result.get('values', [])

    numCols = 0
    for col in values[0]:
        if col != "":
            numCols += 1
        else:
            break
    numItems = numCols - 2
    print("{} clothing items".format(numItems))

    # Get number of rows (number of people)
    range2 = 'C2:C100'
    result = sheet.values().get(spreadsheetId=LULU_ORDER_SHEET,
                                range=range2).execute()
    values = result.get('values', [])

    numPeople = 0
    for row in values:
        if len(row) != 0:
            numPeople += 1
        else:
            break
    print("{} unique orders".format(numPeople))

    itemColumnLetters = {0:"D", 1:"E", 2:"F", 3:"G", 4:"H", 5:"I", 6:"J", 7:"K", 8:"L", 9:"M", 10:"N", 11:"O", 12:"P", 13:"Q", 14:"R", 15:"S", 16:"T", 17:"U", 18:"V", 19:"W", 20:"X", 21:"Y", 22:"Z"}

    for i in range(0, numItems):
        colLetter = itemColumnLetters[i]
        # Create range for current index
        currRange = "{}1:{}{}".format(colLetter, colLetter, numPeople + 1)
        # Get values
        result = sheet.values().get(spreadsheetId=LULU_ORDER_SHEET,
                                range=currRange).execute()
        values = result.get('values', [])

        # Remove and store title
        itemName = values.pop(0)
        # Store all sizes in dictionary with count
        selectedSizes = {}
        for row in values:
            if row[0] not in selectedSizes:
                selectedSizes[row[0]] = 1
            else:
                selectedSizes[row[0]] += 1

        print("->GOT_SIZES: {}".format(selectedSizes))
        
        # Set up data to enter
        values = []
        body = {
            'values': values
        }
        # Remove not interested
        if "Not Interested" in selectedSizes:
            selectedSizes.pop("Not Interested")
        # Add title to values
        values.append(itemName)
        i = 0
        for size, count in selectedSizes.items():
            values.append( ["{} -> {}".format(size, count)] )
				
            i += 1

        # Write Changes
        startRange = 28
        endRange = startRange + len(selectedSizes)
        outputRange = "{}{}:{}{}".format(colLetter, startRange, colLetter, endRange)
        result = sheet.values().update(
        spreadsheetId=LULU_ORDER_SHEET, range=outputRange,
        valueInputOption="USER_ENTERED", body=body).execute()
        selectedSizes = {}
            
    if DELETE:
        print("DELETING")
        clearValues = []
        clearBody = {
            'values': clearValues
        }
        for i in range(0, numItems):
            print("Deleted one ")
            colLetter = itemColumnLetters[i]
            
            # Write Changes
            startRange = 27
            endRange = startRange + 12
             # Set up data to enter
            for _ in range(0, 12):
                clearValues.append( [""] )
            outputRange = "{}{}:{}{}".format(colLetter, startRange, colLetter, endRange)
            print("Range is: {}".format(outputRange))
            result = sheet.values().update(
            spreadsheetId=LULU_ORDER_SHEET, range=outputRange,
            valueInputOption="USER_ENTERED", body=clearBody).execute()
            clearValues = []
            


if __name__ == '__main__':
    main()
