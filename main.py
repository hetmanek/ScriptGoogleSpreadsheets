import os
import math

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Define the required OAuth2 scopes for Google Sheets API
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# Specify the Google Sheets spreadsheet ID and range
SPREADSHEET_ID = "1ejAYsOrIVFp04UlnCB_wQO8d9mdJWJgH3fi51tP2Uzo"
SPREADSHEET_RANGE_NAME = "engenharia_de_software!A1:H27"

# Minimum passing grade and minimum grade to reject a student
MIN_PASSING_GRADE = 70
MIN_REJECT_GRADE = 50


def get_credentials():
    """
    Retrieves the user's credentials for accessing Google Sheets.

    Returns:
        Credentials: The user's credentials.
    """
    credentials = None
    if os.path.exists("token.json"):
        credentials = Credentials.from_authorized_user_file("token.json", SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not credentials or not credentials.valid:
        if credentials and credentials.expired and credentials.refresh_token:
            credentials.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            credentials = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open("token.json", "w") as token:
                token.write(credentials.to_json())
    return credentials


def update_values(values):
    """
    Updates the values in the spreadsheet based on grading criteria.

    Args:
        values (list): The values retrieved from the spreadsheet.

    Returns:
        list: The updated values.
    """
    try:
        # Extract the total number of classes from the second row
        total_number_of_classes = int(''.join(caracter for caracter in values[1][0] if caracter.isdigit()))

        # Iterate over rows starting from the third row
        for row in range(3, len(values)):
            if len(values[row]) < 7:
                values[row].append('')
                values[row].append('')

            # Verify if the current studend have minimal frequency
            if int(values[row][2]) > (0.25 * total_number_of_classes):
                values[row][6] = 'Reprovado por Falta'
                values[row][7] = '0'

            else:
                # Calculate the average of grades for the current student
                grades = [int(grade) for grade in values[row][3:6]]
                average = sum(grades) / len(grades)

                # Apply grading criteria
                if average >= MIN_PASSING_GRADE:
                    values[row][6] = 'Aprovado'
                    values[row][7] = '0'

                elif average < MIN_REJECT_GRADE:
                    values[row][6] = 'Reprovado por Nota'
                    values[row][7] = '0'

                else:
                    grade_for_final_approval = (2 * MIN_REJECT_GRADE) - average
                    values[row][6] = 'Exame Final'
                    values[row][7] = str(math.ceil(grade_for_final_approval))
    except Exception as e:
        print(f"Error during values update: {e}")
        return None
    return values


def main():
    """
    Main function to execute the Google Sheets update process.
    """
    try:
        # Obtain user credentials
        credentials = get_credentials()

        # Build the Google Sheets API service
        google_api_service = build("sheets", "v4", credentials=credentials)

        # Retrieve values from the specified spreadsheet and range
        spreadsheet = google_api_service.spreadsheets().values().get(spreadsheetId=SPREADSHEET_ID,
                                                                     range=SPREADSHEET_RANGE_NAME).execute()

        # Update values based on grading criteria
        updated_values = update_values(spreadsheet.get("values", []))

        if updated_values is not None:
            # Update the spreadsheet with the modified values
            google_api_service.spreadsheets().values().update(spreadsheetId=SPREADSHEET_ID,
                                                              range=SPREADSHEET_RANGE_NAME,
                                                              valueInputOption="USER_ENTERED",
                                                              body={"values": updated_values}).execute()
            print("Values successfully updated.")
    except HttpError as error:
        print(f"Error during execution: {error}")


if __name__ == "__main__":
    main()
