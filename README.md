# SharePoint Sync Tool

Inkrementelles Upload-Tool für SharePoint Online. Lädt Dateien aus einem lokalen Ordner nach SharePoint hoch — bei Wiederholungsläufen werden nur neue oder geänderte Dateien übertragen.

## Features

- **Delta-Sync**: Erkennt Änderungen über mtime+Größe (schnell) → MD5-Hash (zuverlässig)
- **Große Dateien**: Upload-Sessions für Dateien > 4 MB, Chunks à 320 KiB
- **Throttling**: Respektiert `Retry-After` bei HTTP 429/503
- **Optionales Löschen**: Lokal gelöschte Dateien können auch remote entfernt werden
- **Cross-Platform**: Läuft auf Windows, Linux und macOS
- **State-Datei**: Atomares JSON-Update für Wiederholbarkeit

---

## Voraussetzungen

- Python 3.11+
- Eine Entra ID App Registration mit der Berechtigung `Sites.ReadWrite.All` (Application)

---

## Installation

```bash
pip install -r requirements.txt
```

---

## Entra ID Setup

Es gibt drei Wege, die App Registration einzurichten:

### Option A: Terraform (empfohlen)

```bash
cd terraform
az login
terraform init
terraform apply
```

Terraform erstellt die App Registration, erteilt Admin Consent und schreibt die `.env` Datei automatisch.

**Variablen** (`terraform/variables.tf`):

| Variable | Default | Beschreibung |
|---|---|---|
| `app_display_name` | `SharePoint-Uploader` | Name der App Registration |
| `secret_expiry_date` | `2027-01-01T00:00:00Z` | Ablaufdatum des Secrets |
| `write_env_file` | `true` | `.env` automatisch schreiben |

### Option B: Python-Skript

```bash
az login
python setup_entra.py --app-name "SharePoint-Uploader"
```

Das Skript erstellt App Registration, Service Principal und Client Secret und schreibt die `.env` Datei. Voraussetzung: Azure CLI eingeloggt mit Global Admin oder Application Administrator Rechten.

### Option C: Manuell im Azure Portal

1. [portal.azure.com](https://portal.azure.com) → **Entra ID** → **App-Registrierungen** → **Neue Registrierung**
2. Name wählen, "Nur dieser Organisationsverzeichnis" auswählen
3. **API-Berechtigungen** → **Berechtigung hinzufügen** → Microsoft Graph → Anwendungsberechtigungen → `Sites.ReadWrite.All`
4. **Administratorzustimmung erteilen**
5. **Zertifikate & Geheimnisse** → **Neuer geheimer Clientschlüssel**
6. Werte in `.env` eintragen (siehe `.env.example`)

---

## Konfiguration

### `.env` (Secrets — nicht in Git)

```env
AZURE_TENANT_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
AZURE_CLIENT_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
AZURE_CLIENT_SECRET=dein-secret
```

Kopiere `.env.example` zu `.env` und trage die Werte ein.

### `config.yaml`

```yaml
source_folder: "C:/Dokumente/Sync"   # Quellordner (Windows- oder Unix-Pfad)

sharepoint:
  site_url: "https://contoso.sharepoint.com/sites/MySite"
  target_library: "Dokumente"         # Name der Dokumentenbibliothek
  target_folder: "Upload"             # Unterordner in der Bibliothek

sync:
  delete_remote_if_local_deleted: false  # Lokal gelöschte Dateien auch remote löschen
  upload_chunk_size_mb: 10               # Chunk-Größe für große Dateien

state:
  state_file: "sync_state.json"       # Zustandsdatei

logging:
  log_dir: "logs"
  log_level: "INFO"                   # DEBUG, INFO, WARNING, ERROR
```

---

## Verwendung

```bash
# Normaler Sync
python -m sharepoint_sync

# Mit eigener Config
python -m sharepoint_sync --config /pfad/zu/config.yaml

# Nur anzeigen, was hochgeladen würde
python -m sharepoint_sync --dry-run

# Alle Dateien erneut hochladen (State ignorieren)
python -m sharepoint_sync --force-full-sync

# Debug-Logging
python -m sharepoint_sync --log-level DEBUG
```

**Exit Codes:**
- `0` — Erfolgreich
- `2` — Fehler aufgetreten

---

## Wie es funktioniert

```
FolderWalker → ChangeDetector → SharePointClient → StateManager
     ↓               ↓                  ↓               ↓
Alle Dateien   mtime+size →       Upload small      JSON-State
rekursiv       UNCHANGED?         (≤4 MB PUT)       atomar
               Nein → MD5         Upload large      speichern
               UNCHANGED?         (>4 MB Session)
               Nein → Upload
```

1. **FolderWalker** geht rekursiv durch den Quellordner
2. **ChangeDetector** vergleicht mtime+Größe mit dem State — nur bei Abweichung wird MD5 berechnet
3. **SharePointClient** lädt neue/geänderte Dateien hoch und erstellt fehlende Ordner automatisch
4. **StateManager** speichert mtime, Größe, MD5 und SharePoint-Item-ID atomar in JSON

---

## Übersprungene Dateien

| Muster | Beispiel |
|---|---|
| Versteckte Dateien/Ordner | `.git/`, `.env` |
| Office-Temp-Dateien | `~$Dokument.docx` |
| System-Dateien | `Thumbs.db`, `desktop.ini`, `.DS_Store` |

---

## Projektstruktur

```
sharepoint-upload/
├── .env.example           # Template für Credentials
├── .gitignore
├── config.yaml            # Konfiguration
├── requirements.txt
├── setup_entra.py         # Entra ID Setup (Python)
├── terraform/
│   ├── main.tf            # Entra ID Setup (Terraform)
│   ├── variables.tf
│   └── outputs.tf
└── sharepoint_sync/
    ├── __main__.py        # python -m sharepoint_sync
    ├── cli.py             # CLI / argparse
    ├── config.py          # Konfiguration laden
    ├── auth.py            # MSAL Token
    ├── graph_client.py    # HTTP + Retry-Logik
    ├── sharepoint_client.py  # Graph API SP-Operationen
    ├── state_manager.py   # Zustandsdatei (JSON)
    ├── change_detector.py # Delta-Erkennung
    ├── folder_walker.py   # Ordner rekursiv durchlaufen
    ├── uploader.py        # Sync-Loop
    └── logger.py          # Logging Setup
```

---

## Secret erneuern

Wenn das Client Secret abläuft:

**Via Terraform:**
```bash
terraform taint azuread_service_principal_password.sharepoint_sync
terraform apply
```

**Via Python-Skript:**
```bash
python setup_entra.py --app-name "SharePoint-Uploader"
```

**Manuell:** Im Azure Portal → App Registration → Zertifikate & Geheimnisse → Neues Secret erstellen.

---

## Sicherheitshinweise

- Die `.env` Datei enthält sensible Credentials und ist in `.gitignore` ausgeschlossen
- Das Client Secret hat ein Ablaufdatum — nach Ablauf schlägt die Authentifizierung fehl
- `Sites.ReadWrite.All` gibt der App Lese- und Schreibzugriff auf **alle** SharePoint-Sites im Tenant
- Für produktive Umgebungen: Secret in einem Key Vault hinterlegen und via Managed Identity abrufen
