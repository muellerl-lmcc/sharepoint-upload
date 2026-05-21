#!/usr/bin/env python3
"""
Automatisches Entra ID Setup für das SharePoint Sync Tool.

Voraussetzungen:
  - Azure CLI installiert und eingeloggt: az login
  - Der eingeloggte Account hat Rechte, App Registrations zu erstellen
    und Admin Consent zu erteilen (Global Admin oder Application Admin)

Ausführung:
  python setup_entra.py --app-name "SharePoint-Uploader"
"""

import argparse
import json
import subprocess
import sys
import time
from pathlib import Path

MS_GRAPH_APP_ID = "00000003-0000-0000-c000-000000000000"
SITES_READWRITE_ALL_ROLE_ID = "9492366f-7969-46a4-8d15-ed1a20078fff"


def run(cmd: list[str]) -> dict:
    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode != 0:
        print(f"FEHLER: {result.stderr.strip()}", file=sys.stderr)
        sys.exit(1)
    return json.loads(result.stdout) if result.stdout.strip() else {}


def get_az_token() -> str:
    print("Schritt 1/6: Azure CLI Token abrufen...")
    result = subprocess.run(
        ["az", "account", "get-access-token", "--resource", "https://graph.microsoft.com"],
        capture_output=True, text=True,
    )
    if result.returncode != 0:
        print("Fehler: Azure CLI nicht eingeloggt. Bitte zuerst 'az login' ausführen.", file=sys.stderr)
        sys.exit(1)
    return json.loads(result.stdout)["accessToken"]


def graph_post(token: str, path: str, body: dict) -> dict:
    import urllib.request
    url = f"https://graph.microsoft.com/v1.0{path}"
    data = json.dumps(body).encode("utf-8")
    req = urllib.request.Request(url, data=data, method="POST")
    req.add_header("Authorization", f"Bearer {token}")
    req.add_header("Content-Type", "application/json")
    with urllib.request.urlopen(req) as resp:
        return json.loads(resp.read())


def graph_get(token: str, path: str) -> dict:
    import urllib.request
    url = f"https://graph.microsoft.com/v1.0{path}"
    req = urllib.request.Request(url)
    req.add_header("Authorization", f"Bearer {token}")
    with urllib.request.urlopen(req) as resp:
        return json.loads(resp.read())


def main() -> None:
    parser = argparse.ArgumentParser(description="Entra ID App Registration für SharePoint Sync")
    parser.add_argument("--app-name", default="SharePoint-Uploader", help="Name der App Registration")
    parser.add_argument("--env-file", default=".env", help="Pfad zur .env Ausgabedatei (default: .env)")
    args = parser.parse_args()

    token = get_az_token()

    # Tenant ID ermitteln
    account_info = json.loads(subprocess.run(["az", "account", "show"], capture_output=True, text=True).stdout)
    tenant_id = account_info["tenantId"]
    print(f"  Tenant ID: {tenant_id}")

    # Schritt 2: App Registration erstellen
    print(f"Schritt 2/6: App Registration '{args.app_name}' erstellen...")
    app = graph_post(token, "/applications", {
        "displayName": args.app_name,
        "signInAudience": "AzureADMyOrg",
        "requiredResourceAccess": [
            {
                "resourceAppId": MS_GRAPH_APP_ID,
                "resourceAccess": [
                    {"id": SITES_READWRITE_ALL_ROLE_ID, "type": "Role"}
                ],
            }
        ],
    })
    app_id = app["appId"]
    app_object_id = app["id"]
    print(f"  App ID (Client ID): {app_id}")

    # Schritt 3: Service Principal erstellen
    print("Schritt 3/6: Service Principal erstellen...")
    sp = graph_post(token, "/servicePrincipals", {"appId": app_id})
    sp_object_id = sp["id"]
    print(f"  Service Principal Object ID: {sp_object_id}")

    # Schritt 4: Microsoft Graph Service Principal ermitteln
    print("Schritt 4/6: Microsoft Graph Service Principal ermitteln...")
    time.sleep(2)  # kurz warten bis SP angelegt ist
    graph_sps = graph_get(token, f"/servicePrincipals?$filter=appId eq '{MS_GRAPH_APP_ID}'")
    graph_sp_id = graph_sps["value"][0]["id"]

    # Schritt 5: Admin Consent (App Role Assignment)
    print("Schritt 5/6: Admin Consent erteilen (Sites.ReadWrite.All)...")
    graph_post(token, "/servicePrincipals/{}/appRoleAssignments".format(sp_object_id), {
        "principalId": sp_object_id,
        "resourceId": graph_sp_id,
        "appRoleId": SITES_READWRITE_ALL_ROLE_ID,
    })
    print("  Admin Consent erteilt.")

    # Schritt 6: Client Secret erstellen
    print("Schritt 6/6: Client Secret erstellen...")
    secret_resp = graph_post(token, f"/applications/{app_object_id}/addPassword", {
        "passwordCredential": {
            "displayName": "sharepoint-sync-secret",
            "endDateTime": "2027-01-01T00:00:00Z",
        }
    })
    client_secret = secret_resp["secretText"]

    # .env Datei schreiben
    env_path = Path(args.env_file)
    env_content = f"""AZURE_TENANT_ID={tenant_id}
AZURE_CLIENT_ID={app_id}
AZURE_CLIENT_SECRET={client_secret}
"""
    env_path.write_text(env_content, encoding="utf-8")
    print(f"\nFertig! Credentials gespeichert in: {env_path.resolve()}")
    print(f"  Tenant ID:     {tenant_id}")
    print(f"  Client ID:     {app_id}")
    print(f"  Client Secret: {'*' * 8}{client_secret[-4:]}")
    print("\nHINWEIS: Das Client Secret ist nur einmal sichtbar. Bewahre die .env Datei sicher auf!")


if __name__ == "__main__":
    main()
