#!/usr/bin/env python3
"""
Prüft ob Konfiguration und Azure-Setup vollständig und funktionsfähig sind.

Ausführung:
  python check_setup.py
  python check_setup.py --config config.yaml
"""

import argparse
import os
import sys

OK = "\033[32m✓\033[0m"
FAIL = "\033[31m✗\033[0m"
WARN = "\033[33m!\033[0m"


def check(label: str, ok: bool, detail: str = "") -> bool:
    symbol = OK if ok else FAIL
    line = f"  {symbol}  {label}"
    if detail:
        line += f"  ({detail})"
    print(line)
    return ok


def section(title: str) -> None:
    print(f"\n{title}")
    print("─" * len(title))


def main() -> None:
    parser = argparse.ArgumentParser(description="Prüft Konfiguration und Azure-Setup")
    parser.add_argument("--config", default="config.yaml", help="Pfad zur config.yaml (default: config.yaml)")
    args = parser.parse_args()

    all_ok = True

    # ── 1. .env ──────────────────────────────────────────────────────────────
    section("1. Umgebungsvariablen (.env)")

    env_file_exists = os.path.exists(".env")
    all_ok &= check(".env Datei vorhanden", env_file_exists, ".env.example als Vorlage verwenden" if not env_file_exists else "")

    if env_file_exists:
        from dotenv import load_dotenv
        load_dotenv()

    required_env = ["AZURE_TENANT_ID", "AZURE_CLIENT_ID", "AZURE_CLIENT_SECRET"]
    for var in required_env:
        val = os.environ.get(var, "")
        all_ok &= check(f"{var} gesetzt", bool(val), "leer oder nicht gesetzt" if not val else f"{val[:4]}****")

    # ── 2. config.yaml ───────────────────────────────────────────────────────
    section("2. config.yaml")

    config_exists = os.path.exists(args.config)
    all_ok &= check(f"{args.config} vorhanden", config_exists)

    config = None
    if config_exists:
        try:
            import yaml
            with open(args.config, "r", encoding="utf-8") as f:
                raw = yaml.safe_load(f)

            source = raw.get("source_folder", "")
            all_ok &= check("source_folder gesetzt", bool(source), source or "leer")

            sp = raw.get("sharepoint", {})
            site_url = sp.get("site_url", "")
            placeholder = "contoso.sharepoint.com" in site_url
            all_ok &= check("sharepoint.site_url gesetzt", bool(site_url) and not placeholder,
                            "noch Beispielwert" if placeholder else (site_url or "leer"))
            all_ok &= check("sharepoint.target_library gesetzt", bool(sp.get("target_library")),
                            sp.get("target_library", "leer"))
            all_ok &= check("sharepoint.target_folder gesetzt", bool(sp.get("target_folder")),
                            sp.get("target_folder", "leer"))

            # Load full config for later checks
            from sharepoint_sync.config import load_config
            config = load_config(args.config)

        except Exception as e:
            all_ok &= check("config.yaml parsebar", False, str(e))

    # ── 3. Quellordner ───────────────────────────────────────────────────────
    section("3. Quellordner")

    if config:
        source_exists = os.path.isdir(config.source_folder)
        all_ok &= check(f"Ordner existiert: {config.source_folder}", source_exists,
                        "Pfad nicht gefunden" if not source_exists else "")
    else:
        print(f"  {WARN}  Übersprungen (config.yaml nicht ladbar)")

    # ── 4. Azure Token ───────────────────────────────────────────────────────
    section("4. Azure AD / Entra ID — Token")

    if config:
        try:
            from sharepoint_sync.auth import SharePointAuthenticator
            auth = SharePointAuthenticator(config.tenant_id, config.client_id, config.client_secret)
            token = auth.get_token()
            all_ok &= check("Token-Abruf erfolgreich (MSAL)", bool(token),
                            f"Token-Länge: {len(token)} Zeichen")
        except Exception as e:
            all_ok &= check("Token-Abruf erfolgreich (MSAL)", False, str(e))
    else:
        print(f"  {WARN}  Übersprungen (Konfiguration unvollständig)")

    # ── 5. SharePoint-Verbindung ─────────────────────────────────────────────
    section("5. SharePoint-Verbindung")

    if config:
        try:
            from sharepoint_sync.auth import SharePointAuthenticator
            from sharepoint_sync.graph_client import GraphClient
            from sharepoint_sync.sharepoint_client import SharePointClient

            auth = SharePointAuthenticator(config.tenant_id, config.client_id, config.client_secret)
            graph = GraphClient(auth.get_token, max_retries=2, backoff_base=1)
            sp = SharePointClient(graph, config.sharepoint.site_url, config.sharepoint.target_library,
                                  config.sharepoint.target_folder)

            site_id = sp._get_site_id()
            all_ok &= check(f"Site erreichbar: {config.sharepoint.site_url}", True, f"ID: {site_id[:16]}…")

            drive_id = sp._get_drive_id()
            all_ok &= check(f"Bibliothek gefunden: {config.sharepoint.target_library}", True,
                            f"Drive ID: {drive_id[:16]}…")

        except Exception as e:
            all_ok &= check("SharePoint-Verbindung", False, str(e))
    else:
        print(f"  {WARN}  Übersprungen (Konfiguration unvollständig)")

    # ── Ergebnis ─────────────────────────────────────────────────────────────
    print()
    if all_ok:
        print(f"{OK} Alle Checks bestanden — Setup ist vollständig und einsatzbereit.\n")
        sys.exit(0)
    else:
        print(f"{FAIL} Einige Checks fehlgeschlagen — bitte Konfiguration prüfen.\n")
        sys.exit(1)


if __name__ == "__main__":
    main()
