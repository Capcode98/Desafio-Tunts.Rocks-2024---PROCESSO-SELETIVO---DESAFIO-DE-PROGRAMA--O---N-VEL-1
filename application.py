import os.path
import math

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = "1kYakLPO_0ydlYg98ciKZlEbSuFwWKyVsMsuG5yNCRW8"
SAMPLE_RANGE_NAME = "engenharia_de_software!A2:H27"

# Validation of the credentials file (token.json)
def Validation():

  print("Validating credentials...")

  """Shows basic usage of the Sheets API.
  Prints values from a sample spreadsheet.
  """

  creds = None

  # The file token.json stores the user's access and refresh tokens, and is
  # created automatically when the authorization flow completes for the first
  # time.

  if os.path.exists("token.json"):

    creds = Credentials.from_authorized_user_file("token.json", SCOPES)

  # If there are no (valid) credentials available, let the user log in.

  if not creds or not creds.valid:

    if creds and creds.expired and creds.refresh_token:

      creds.refresh(Request())

    else:

      flow = InstalledAppFlow.from_client_secrets_file("client_secret.json", SCOPES)

      creds = flow.run_local_server(port=0)

    # Save the credentials for the next run

    with open("token.json", "w") as token:
      token.write(creds.to_json())

      print("Credentials validated successfully!")
    
  return creds

# Get the total number of classes per semester
def Get_Total_Classes_Per_Semester(creds):

  print("Getting the total number of classes per semester...")
    
  service = build("sheets", "v4", credentials=creds)

  # Call the Sheets API
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

  total_classes = values[0][0]

  print("Total number of classes per semester retrieved successfully!")

  return total_classes.split(": ")[1]

# Get the additional information from the spreadsheet
def Get_Additional_Information(creds, SAMPLE_RANGE_NAME):

  print("Getting additional information from the spreadsheet...")

  service = build("sheets", "v4", credentials=creds)

  # Call the Sheets API
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
  
  else:
    print("Additional information retrieved successfully!")

    return values

# Calculate the student status 
def Calculate_Student_Status(final_average, absences, total_classes):

  if absences > (int(total_classes)/4):
    return "Reprovado por Falta"

  if final_average < 5:
    return "Reprovado por Nota"
  
  elif final_average >= 5 and final_average < 7: 
    return f"Exame Final"
  
  elif final_average >= 7:
    return "Aprovado"

# Calculate the grade for final approval
def Calculate_Grade_For_Final_Approval(final_average,absences, total_classes):

  if absences > (int(total_classes)/4):
    return "0"

  if final_average < 5:
    return "0"
  
  elif final_average >= 5 and final_average < 7: 
    grade_for_final_approval = 10 - final_average 
    return str(grade_for_final_approval)
  
  elif final_average >= 7:
    return "0"

# Insert the values into the spreadsheet
def Insert_Values(creds, values, SAMPLE_RANGE_NAME):

  print("Inserting the values into the spreadsheet...")

  service = build("sheets", "v4", credentials=creds)
      
  body = {"values": values}

  print("Values inserted successfully!")

  return service.spreadsheets().values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID,range=SAMPLE_RANGE_NAME,valueInputOption="USER_ENTERED",body=body,).execute()


def main():

  creds = Validation()
  
  try:

    values_to_insert = []

    total_classes = Get_Total_Classes_Per_Semester(creds=creds)

    values = Get_Additional_Information(creds=creds, SAMPLE_RANGE_NAME="engenharia_de_software!A4:H27")

    print(values)

    print("Calculating the student status...And the grade for final approval...")

    # Print the values from the spreadsheet to check that everything is working properly.
    for row in values:
      
      id = row[0]
      name = row[1]
      absences = int(row[2])
      grade1 = int(row[3])
      grade2 = int(row[4])
      grade3 = int(row[5])
      final_average = float((grade1 + grade2 + grade3)/3)
      final_average = math.ceil(final_average)/10
      
      values_to_insert.append([Calculate_Student_Status(final_average, absences, total_classes), Calculate_Grade_For_Final_Approval(final_average, absences, total_classes)])

      print(f'ID: {id}, name: {name}, absences: {absences}, grade1: {grade1}, grade2: {grade2}, grade3: {grade3}, final_average: {final_average}, Situation: {Calculate_Student_Status(final_average, absences, total_classes)}\n')
    
    Insert_Values(creds=creds, values=values_to_insert, SAMPLE_RANGE_NAME="engenharia_de_software!G4")

  except HttpError as err:
    print(err)


if __name__ == "__main__":
  main()