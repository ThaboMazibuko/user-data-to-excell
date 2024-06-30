import openpyxl
import re

def is_valid_email(email):
    # Simple regex for validating an email
    regex = r'^\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
    return re.match(regex, email)

def is_valid_phone(phone):
    # Regex for validating a phone number (starts with 0, exactly 10 digits)
    regex = r'^0\d{9}$'
    return bool(re.match(regex, phone))

# Create a new workbook and select the active worksheet
wb = openpyxl.Workbook()
ws = wb.active

# Set the headers in the first row
ws.append(["Name", "Surname", "Email", "Cellphone Number"])

while True:
    try:
        # Request user input with validation
        name = input("Enter your name: ").strip()
        if not name:
            raise ValueError("Name cannot be empty.")

        surname = input("Enter your surname: ").strip()
        if not surname:
            raise ValueError("Surname cannot be empty.")

        email = input("Enter your email: ").strip()
        if not is_valid_email(email):
            raise ValueError("Invalid email format, use format: email@gmail.com")

        cellphone = input("Enter your cellphone number: ").strip()
        if not is_valid_phone(cellphone):
            raise ValueError("Invalid cellphone number format, use format: 0789654878")
        
        # Append the input data to the worksheet
        ws.append([name, surname, email, cellphone])

        # Ask if the user wants to add another entry
        continue_input = input("Do you want to add another person? (yes/no): ").strip().lower()
        if continue_input != 'yes':
            break

    except ValueError as e:
        print(f"Error: {e}. Please try again.")

# Save the workbook
file_name = "user_data.xlsx"
wb.save(file_name)

print(f"Data saved to {file_name}")
