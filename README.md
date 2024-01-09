# Product-Registry-System
This script automated part of a mid-sized company's internal product registry system.

## Functionalities:
1. Loops through the responses of a Google Forms Responses sheet, checking to see if an email has been sent yet
2. Cross-references the new submission's device id with another spreadsheet, grabbing the corresponding device password from the second sheet
3. Sends an email with the password to the user's submitted email, and sends an error email if the id is invalid
