import gspread
import smtplib
from google.oauth2 import service_account
from email.mime.text import MIMEText

# authenitcate Sheets API
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
creds = service_account.Credentials.from_service_account_file('json_file_path', scopes=SCOPES)
client = gspread.authorize(creds)

# open the response and password spreadsheets
responseSheet = client.open_by_key('responses_sheet_key').sheet1
passwordSheet = client.open_by_key('passwords_sheet_key').worksheet('Device Passwords')

responseValues = responseSheet.get_all_values() # all values as list of lists
passwordValues = passwordSheet.col_values(1) # all values in first column as list of strings

# format message, login, and send message
def sendEmail(text, subject, emailAddress):
    msg = MIMEText(text, 'html')
    msg['Subject'] = subject
    msg['To'] = emailAddress
    msg['From'] = "our@email.com"
    smtp_server = smtplib.SMTP_SSL('smtp.gmail.com', 465) #SMTP server connection
    smtp_server.login("our@email.com", "app_password")
    smtp_server.sendmail("our@email.com", emailAddress, msg.as_string())
    smtp_server.quit()
    responseSheet.update(f'G{responseValues.index(row)+1}', 'done')
    print(F'sent message to {emailAddress}')
    
# loop through each row in the response sheet
for row in responseValues[1:]:  # skip the header row
    name = row[1]
    email = row[2]
    productID = row[3]
    confirmation = row[6]

    if confirmation != "done" :
        
        if email and productID not in passwordValues:
            print('id not found')
            message=(f"""<html>
                    <head></head>
                    <body>
                        Hello {name},<br /><br />
                        
                        We can't find the <b>Device ID ({productID})</b> you submitted in the Product Registration form. <br /><br />
                        
                        Please double check that you submitted the ID correctly. <br /><br />
                        
                        If this issue persists, please contact us at itemailgroup@email.com.<br /><br />
                        
                        Thank you,<br />
                        IT Asset Team
                    </body>
                </html>
            """)
            
            sendEmail(message, "Product Registration (ERROR)", email)
        elif email and productID in passwordValues :
        # crossreference the responses sheet with the password sheet with the product id
            passwordRow = passwordSheet.find(productID).row
            password = passwordSheet.cell(passwordRow, 3).value

            # compose message
            message=(f"""\
                <html>
                    <head></head>
                    <body>
                        Hello {name},<br /><br />
                        
                        Thank you for registering your company device.<br /><br />
                        
                        As part of the registration process, please find your device password below:<br /><br />
                        
                        <b>Device password: {password}</b><br /><br />
                        
                        Please use this password to log into your new device. Once you have logged in, we recommend changing your device password to ensure the security of your device and data.<br /><br />
                        
                        Should you encounter any ssues with your device, please don't hesitate to contact us at itemailgroup@email.com and our team will be happy to assist you.<br /><br />
                        
                        Once again, thank you for your cooperation, and we look forward to supporting you and your device needs.<br /><br />
                        
                        Thank you,<br />
                        IT Asset Team
                    </body>
                </html>
            """)
            sendEmail(message, f"Product Registration ({productID})", email)
        else :
            print("emails up to date")
            break