import os
import requests
from dotenv import load_dotenv
import msal

load_dotenv()

CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
TENANT_ID = os.getenv("TENANT_ID")

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPE = ["https://graph.microsoft.com/.default"]


class OutlookClient:
    def __init__(self):
        self.app = msal.ConfidentialClientApplication(
            CLIENT_ID,
            authority=AUTHORITY,
            client_credential=CLIENT_SECRET
        )

    def get_access_token(self):
        result = self.app.acquire_token_for_client(scopes=SCOPE)

        if "access_token" in result:
            return result["access_token"]
        else:
            raise Exception(f"Token error: {result}")

    def get_emails(self, top=10):
        token = self.get_access_token()

        headers = {
            "Authorization": f"Bearer {token}"
        }

        url = f"https://graph.microsoft.com/v1.0/me/messages?$top={top}"

        response = requests.get(url, headers=headers)

        if response.status_code != 200:
            raise Exception(response.text)

        return response.json().get("value", [])




if __name__ == "__main__":
    client = OutlookClient()
    emails = client.get_emails()

    for email in emails:
        print("Subject:", email.get("subject"))
        print("From:", email.get("from", {}).get("emailAddress", {}).get("address"))
        print("-" * 40)