"""
generate_token.py
-----------------
Run this once to authenticate with Google and generate token.pickle.
After running, token.pickle will be created in the project root.
token.pickle is gitignored — never commit it.

Usage:
    python generate_token.py

Requirements:
    - client_secret.json must exist in the project root (copied from client_secret_template.json
      and filled with your real GCP credentials).
    - The Google account you authenticate with must have edit access to the target Google Sheet.
"""

import pickle
from pathlib import Path

from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

SCOPES = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive",
]

BASE_DIR = Path(__file__).resolve().parent
CLIENT_SECRET_PATH = BASE_DIR / "client_secret.json"
TOKEN_PATH = BASE_DIR / "token.pickle"


def generate_token():
    creds = None

    if TOKEN_PATH.exists():
        with TOKEN_PATH.open("rb") as f:
            creds = pickle.load(f)

    if creds and creds.valid:
        print("Existing token is still valid. No action needed.")
        return

    if creds and creds.expired and creds.refresh_token:
        print("Refreshing expired token...")
        creds.refresh(Request())
    else:
        if not CLIENT_SECRET_PATH.exists():
            raise FileNotFoundError(
                f"client_secret.json not found at {CLIENT_SECRET_PATH}.\n"
                "Copy client_secret_template.json, rename it to client_secret.json, "
                "and fill in your GCP credentials."
            )
        print("Starting OAuth flow. A browser window will open for authentication...")
        flow = InstalledAppFlow.from_client_secrets_file(str(CLIENT_SECRET_PATH), SCOPES)
        creds = flow.run_local_server(port=0)

    with TOKEN_PATH.open("wb") as f:
        pickle.dump(creds, f)

    print(f"Token saved to {TOKEN_PATH}")
    print("You can now run the Streamlit app.")


if __name__ == "__main__":
    generate_token()
