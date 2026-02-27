import msal
import requests
from pathlib import Path

AUTHORITY = "https://login.microsoftonline.com/common"
SCOPES = ["Files.Read", "Files.ReadWrite"]
GRAPH_BASE = "https://graph.microsoft.com/v1.0"
TOKEN_CACHE_PATH = Path.home() / ".sarah_editor_token_cache.json"


class OneDriveClient:
    def __init__(self, client_id: str):
        self.client_id = client_id
        self._cache = msal.SerializableTokenCache()
        if TOKEN_CACHE_PATH.exists():
            self._cache.deserialize(TOKEN_CACHE_PATH.read_text())
        self._app = msal.PublicClientApplication(
            client_id, authority=AUTHORITY, token_cache=self._cache
        )
        self._token = None

    def _save_cache(self):
        if self._cache.has_state_changed:
            TOKEN_CACHE_PATH.write_text(self._cache.serialize())

    def authenticate(self):
        """Sign in silently if a cached token exists, otherwise open a browser."""
        accounts = self._app.get_accounts()
        result = None
        if accounts:
            result = self._app.acquire_token_silent(SCOPES, account=accounts[0])
        if not result:
            result = self._app.acquire_token_interactive(scopes=SCOPES)
        self._save_cache()
        if "access_token" not in result:
            raise RuntimeError(
                f"Authentication failed: {result.get('error_description', 'Unknown error')}"
            )
        self._token = result["access_token"]

    def _headers(self):
        return {"Authorization": f"Bearer {self._token}"}

    def list_folder(self, item_id: str = None):
        """Return the children of a folder (or root if item_id is None)."""
        if item_id:
            url = f"{GRAPH_BASE}/me/drive/items/{item_id}/children?$top=500"
        else:
            url = f"{GRAPH_BASE}/me/drive/root/children?$top=500"
        r = requests.get(url, headers=self._headers())
        r.raise_for_status()
        return r.json().get("value", [])

    def get_download_url(self, item_id: str) -> str:
        """Return a short-lived direct download URL for a file."""
        url = f"{GRAPH_BASE}/me/drive/items/{item_id}"
        r = requests.get(url, headers=self._headers())
        r.raise_for_status()
        return r.json().get("@microsoft.graph.downloadUrl", "")

    def create_folder(self, parent_id: str, name: str) -> str:
        """Create a folder inside parent_id and return the new folder's item ID."""
        if parent_id:
            url = f"{GRAPH_BASE}/me/drive/items/{parent_id}/children"
        else:
            url = f"{GRAPH_BASE}/me/drive/root/children"
        r = requests.post(
            url,
            headers={**self._headers(), "Content-Type": "application/json"},
            json={"name": name, "folder": {}, "@microsoft.graph.conflictBehavior": "rename"},
        )
        r.raise_for_status()
        return r.json()["id"]

    def upload_file(
        self,
        parent_id: str,
        filename: str,
        file_path: Path,
        progress_callback=None,
    ):
        """
        Upload a file using a resumable upload session (supports files of any size).
        progress_callback(bytes_uploaded, total_bytes) is called after each chunk.
        """
        if parent_id:
            session_url = (
                f"{GRAPH_BASE}/me/drive/items/{parent_id}:/{filename}:/createUploadSession"
            )
        else:
            session_url = f"{GRAPH_BASE}/me/drive/root:/{filename}:/createUploadSession"

        r = requests.post(
            session_url,
            headers={**self._headers(), "Content-Type": "application/json"},
            json={"item": {"@microsoft.graph.conflictBehavior": "rename"}},
        )
        r.raise_for_status()
        upload_url = r.json()["uploadUrl"]

        file_size = file_path.stat().st_size
        chunk_size = 10 * 1024 * 1024  # 10 MB chunks
        uploaded = 0

        with open(file_path, "rb") as f:
            while uploaded < file_size:
                chunk = f.read(chunk_size)
                end = uploaded + len(chunk) - 1
                r = requests.put(
                    upload_url,
                    data=chunk,
                    headers={
                        "Content-Range": f"bytes {uploaded}-{end}/{file_size}",
                        "Content-Length": str(len(chunk)),
                    },
                )
                if r.status_code not in (200, 201, 202):
                    raise RuntimeError(f"Upload failed at byte {uploaded}: {r.text}")
                uploaded += len(chunk)
                if progress_callback:
                    progress_callback(uploaded, file_size)
