from app.api.outlook_client import OutlookClient

client = OutlookClient()
client.authenticate()

emails = client.get_emails()

for email in emails:
    print(email["subject"])