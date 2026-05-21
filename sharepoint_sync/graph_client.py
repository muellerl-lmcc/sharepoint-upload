import logging
import time
from typing import Any, Callable

import requests

log = logging.getLogger(__name__)

GRAPH_BASE = "https://graph.microsoft.com/v1.0"


class GraphApiError(Exception):
    def __init__(self, status_code: int, message: str) -> None:
        super().__init__(f"HTTP {status_code}: {message}")
        self.status_code = status_code


class GraphNotFoundError(GraphApiError):
    pass


class GraphAuthError(GraphApiError):
    pass


class GraphThrottleError(GraphApiError):
    pass


class GraphClient:
    def __init__(self, token_fn: Callable[[], str], max_retries: int = 5, backoff_base: int = 2) -> None:
        self._token_fn = token_fn
        self._max_retries = max_retries
        self._backoff_base = backoff_base
        self._session = requests.Session()

    def _headers(self) -> dict:
        return {"Authorization": f"Bearer {self._token_fn()}", "Accept": "application/json"}

    def _handle_response(self, resp: requests.Response) -> Any:
        if resp.status_code == 404:
            raise GraphNotFoundError(404, resp.text)
        if resp.status_code in (401, 403):
            raise GraphAuthError(resp.status_code, resp.text)
        if resp.status_code in (429, 503):
            raise GraphThrottleError(resp.status_code, resp.text)
        if not resp.ok:
            raise GraphApiError(resp.status_code, resp.text)
        if resp.status_code == 204 or not resp.content:
            return None
        return resp.json()

    def _request_with_retry(self, method: str, url: str, **kwargs) -> Any:
        for attempt in range(self._max_retries):
            try:
                resp = self._session.request(method, url, **kwargs)
                if resp.status_code in (429, 503):
                    retry_after = int(resp.headers.get("Retry-After", self._backoff_base ** attempt))
                    log.warning("Throttled (HTTP %s), waiting %ss", resp.status_code, retry_after)
                    time.sleep(retry_after)
                    continue
                return self._handle_response(resp)
            except GraphThrottleError as e:
                if attempt == self._max_retries - 1:
                    raise
                wait = self._backoff_base ** attempt
                log.warning("%s — retry %d/%d in %ds", e, attempt + 1, self._max_retries, wait)
                time.sleep(wait)
        raise GraphApiError(429, "Max retries exceeded")

    def get(self, path: str, **kwargs) -> Any:
        url = path if path.startswith("http") else f"{GRAPH_BASE}{path}"
        return self._request_with_retry("GET", url, headers=self._headers(), **kwargs)

    def post(self, path: str, json: Any = None, **kwargs) -> Any:
        url = path if path.startswith("http") else f"{GRAPH_BASE}{path}"
        return self._request_with_retry("POST", url, headers=self._headers(), json=json, **kwargs)

    def put(self, path: str, json: Any = None, **kwargs) -> Any:
        url = path if path.startswith("http") else f"{GRAPH_BASE}{path}"
        return self._request_with_retry("PUT", url, headers=self._headers(), json=json, **kwargs)

    def put_raw(self, url: str, data: bytes, content_range: str, content_length: int) -> Any:
        headers = {
            "Content-Range": content_range,
            "Content-Length": str(content_length),
            "Content-Type": "application/octet-stream",
        }
        resp = self._session.put(url, headers=headers, data=data)
        if resp.status_code in (200, 201, 202):
            return resp.json() if resp.content else None
        raise GraphApiError(resp.status_code, resp.text)

    def delete(self, path: str, **kwargs) -> None:
        url = path if path.startswith("http") else f"{GRAPH_BASE}{path}"
        self._request_with_retry("DELETE", url, headers=self._headers(), **kwargs)
