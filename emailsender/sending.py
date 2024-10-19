import pandas as pd
import yagmail

# Load Excel data
df = pd.read_excel('/Users/admin/PycharmProjects/emailsender/Book1.xlsx', engine='openpyxl')

# Create a status column if it doesn't exist
if 'Status' not in df.columns:
    df['Status'] = ''

# Configure yagmail
yag = yagmail.SMTP('dhamotharan2107@gmail.com', 'nhvdnfnrqxgenbgq')  # Use your app password here

# Loop through each row in the DataFrame
for index, row in df.iterrows():
    name = row['Name']
    email = row['Email']
    subject = row['Subject']
    message = row['Message']

    # Check if email is provided
    if pd.isna(email) or email.strip() == "":
        df.at[index, 'Status'] = 'Fail: No email provided'
        continue

    try:
        # Send the email
        yag.send(
            to=email,
            subject=subject,
            contents=f"Dear {name},\n\n{message}\n\nBest Regards,\nYour Name"
        )
        df.at[index, 'Status'] = 'Success'  # Update status to 'Success' if successful
        print(f'Email sent to {name} at {email} with subject: "{subject}"')
    except Exception as e:
        df.at[index, 'Status'] = 'Fail'  # Update status to 'Fail' if there's an error
        print(f'Failed to send email to {name} at {email}: {str(e)}')

# Save the updated DataFrame back to the Excel file
df.to_excel('/Users/admin/PycharmProjects/emailsender/Book1.xlsx', index=False)

print("All emails processed!")
