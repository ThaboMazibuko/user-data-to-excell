# user-data-to-excell
This Python script collects user information (name, surname, email, and phone) and writes it to an Excel file using openpyxl. It validates the email and phone formats, handles errors, and prompts for multiple entries. Data is saved in user_data.xlsx. The script ensures valid input and robust data collection.
This Python script collects user information (name, surname, email, and cellphone number) and writes it to an Excel file using the openpyxl library. The script performs the following steps:

Import Libraries: Imports openpyxl for handling Excel files and re for regular expressions to validate email and phone formats.
Validation Functions: Defines functions to validate email and phone formats using regular expressions.
Create Workbook: Creates a new Excel workbook and sets up the header row.
User Input Loop: Continuously prompts the user for their name, surname, email, and cellphone number:
Ensures the name and surname are not empty.
Validates the email format.
Validates the phone number format.
Error Handling: Catches any validation errors and prompts the user to re-enter the data.
Continue Prompt: Asks the user if they want to add another entry. If not, it breaks the loop.
Save Workbook: Saves the collected data to an Excel file (user_data.xlsx).