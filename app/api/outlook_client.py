import msal
import requests


class OutlookClient:
    def __init__(self):
        self.client_id = "52cef69b-e50a-40ef-a4aa-34e1fceb209e"
        self.tenant_id = "12fd9da4-2a17-4ab0-af0d-e5e5e3af03b9"

        self.authority = f"https://login.microsoftonline.com/{self.tenant_id}"
        self.scopes = ["Mail.Read"]

        self.token = None

        self.app = msal.PublicClientApplication(
            self.client_id,
            authority=self.authority
        )

    def authenticate(self):
        flow = self.app.initiate_device_flow(scopes=self.scopes)

        if "user_code" not in flow:
            raise Exception(f"Device flow failed: {flow}")

        print("\n=== LOGIN REQUIRED ===")
        print(flow["message"])

        result = self.app.acquire_token_by_device_flow(flow)

        if "access_token" not in result:
            raise Exception(f"Auth failed: {result}")

        self.token = result["access_token"]
        print("✅ Logged in successfully!")

    def get_emails(self, top=5):
        if not self.token:
            raise Exception("Not authenticated")

        headers = {
            "Authorization": f"Bearer {self.token}"
        }

        response = requests.get(
            f"https://graph.microsoft.com/v1.0/me/messages?$top={top}",
            headers=headers
        )

        if response.status_code != 200:
            print(response.text)
            raise Exception("API error")

        return response.json().get("value", [])
        