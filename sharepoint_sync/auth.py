import logging

import msal

log = logging.getLogger(__name__)

GRAPH_SCOPE = ["https://graph.microsoft.com/.default"]


class SharePointAuthenticator:
    def __init__(self, tenant_id: str, client_id: str, client_secret: str) -> None:
        authority = f"https://login.microsoftonline.com/{tenant_id}"
        self._app = msal.ConfidentialClientApplication(
            client_id=client_id,
            client_credential=client_secret,
            authority=authority,
        )

    def get_token(self) -> str:
        result = self._app.acquire_token_for_client(scopes=GRAPH_SCOPE)
        if "access_token" not in result:
            error = result.get("error_description", result.get("error", "unknown"))
            raise RuntimeError(f"Failed to acquire token: {error}")
        log.debug("Token acquired successfully")
        return result["access_token"]
