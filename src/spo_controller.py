"""
Module to manage interaction with SharePoint Online.
"""
import json
import logging
import os

from datetime import datetime
from typing import Optional

import msal
import requests


CHUNK_SIZE: int = 10485760
DEFAULT_TIMEOUT: int = 10
MAX_BACKUPS: int = 4

logger: logging.Logger = logging.getLogger()


class SpoController:
    """
    Class to interface with SharePoint Online and local system secrets related
    to SharePoint Online.

    Attributes:
        authority (str): Authority for Azure AD.
        endpoint (str): Base SharePoint Online endpoint to target.
        scope (str): Scope of access requested in Azure AD.
        client_id (str): Client ID for this app in Azure AD.
        secret (str): Client secret for this app in Azure AD.
        graph_token (str): Dictionary containing the token when authenticating to
                            Azure AD.
        client (ConfidentialClientApplication): MSAL client application.
    """

    def __init__(self) -> None:
        self.authority: str = os.environ.get("SPOBKP_AUTHORITY", "")
        self.endpoint: str = os.environ.get("SPOBKP_ENDPOINT", "")
        self.scope: list = [os.environ.get("SPOBKP_SCOPE", "")]
        self.client_id: str = os.environ.get("SPOBKP_CLIENTID", "")
        self.secret: str = os.environ.get("SPOBKP_SECRET", "")
        self.graph_token: dict = {}
        self.client: Optional[msal.ConfidentialClientApplication] = None

    def validate_info(self) -> bool:
        """
        Validates that all required secrets are available from the OS.

        Returns:
            bool: Whether or not all required properties are found.
        """
        for prop in dir(self):
            if not prop:
                logger.error(
                    "Missing information for %s property. Aborting SharePoint upload.",
                    prop,
                )
                return False

        return True

    def initialize_client(self) -> None:
        """
        Initializes a new client object.
        """
        logger.info("Instantiating an MSAL client.")
        logger.debug("Using Client ID: %s", self.client_id)
        logger.debug("Using Authority: %s", self.authority)
        self.client = msal.ConfidentialClientApplication(
            client_id=self.client_id,
            authority=self.authority,
            client_credential=self.secret,
        )

    def connect_graph(self) -> None:
        """
        Connects to the Microsoft Graph in order to retrieve a token for future
        communication.
        """
        self.initialize_client()
        if not self.client:
            logger.error("MSAL client could not be instantiated.")
            return

        logger.info("Attempting to retrieve an MS Graph token.")
        logger.debug("Using Scope: %s", self.scope)
        result: Optional[dict] = self.client.acquire_token_silent(
            scopes=self.scope, account=None
        )
        if result:
            logger.debug("MSAL token already found. No new attempt required.")
            self.graph_token = result
        else:
            logger.debug("Attempting to retrieve a new MSAL token.")
            result: Optional[dict] = self.client.acquire_token_for_client(
                scopes=self.scope
            )
            if type(result) == dict and "error" in result:
                logger.error(
                    "Received an error retrieving an MSAL token: %s", result["error"]
                )
            else:
                logger.debug("Graph token value: %s", result)
                if type(result) == dict:
                    self.graph_token = result

    def check_token(self) -> bool:
        """
        Ensures that there's a valid token.

        Returns:
            bool: Whether or not a token is found.
        """
        if "access_token" in self.graph_token:
            return True
        return False

    def query_graph(self, endpoint: str = "") -> dict:
        """
        Query the Microsoft Graph at the given endpoint.

        Args:
            endpoint (str): Endpoint to query.

        Returns:
            dict: Results of the query.
        """
        if not endpoint:
            endpoint = self.endpoint
        logger.debug("Attempting a connection to: %s", endpoint)
        resp: requests.Response = requests.get(
            endpoint,
            headers={
                "Authorization": f"Bearer {self.graph_token['access_token']}",
                "Content-Type": "application/json",
            },
            timeout=DEFAULT_TIMEOUT,
        )

        if not resp.status_code == 200:
            logger.error(
                "Graph connection failed with status code: %s", resp.status_code
            )
            return {"error": resp.text, "status": resp.status_code}

        return resp.json()

    def check_dir(self, name: str, endpoint: str = "") -> dict:
        """
        Checks if a directory with the given name exists at the given endpoint.
        Defaults to checking the endpoint value on this instance of the controller
        if nothing is passed.

        Args:
            name (str): Directory name to check for.
            endpoint (str): Endpoint to query for directories.

        Returns:
            dict: Containing the current state.
        """
        if not endpoint:
            endpoint = self.endpoint
        raw_resp: dict = self.query_graph()

        if "error" in raw_resp.keys():
            return raw_resp

        for single_dir in raw_resp.get("value", []):
            if single_dir.get("name", "").lower() == name.lower():
                return {"exists": True}
        return {"exists": False}

    def create_dir(self, name: str, endpoint: str = "") -> bool:
        """
        Creates a directory at the target endpoint.

        Args:
            name (str): Name of the new directory.
            endpoint (str): Where the directory should be created.

        Returns:
            bool: Whether or not the directory was successfully created.
        """
        if not endpoint:
            endpoint = self.endpoint

        payload: dict = {
            "name": name,
            "folder": {},
            "@microsoft.graph.conflictBehavior": "rename",
        }
        accept: str = "application/json;odata.metadata=minimal;odata.streaming=true;"
        accept += "IEEE754Compatible=false;charset=utf-8"
        resp: requests.Response = requests.post(
            endpoint,
            headers={
                "Accept": accept,
                "Authorization": f"Bearer {self.graph_token['access_token']}",
                "Content-Type": "application/json",
            },
            data=json.dumps(payload),
            timeout=DEFAULT_TIMEOUT,
        )

        acceptable_codes: list = [200, 201, 202]
        if not resp.status_code in acceptable_codes:
            logger.error(
                "Failed to create directory %s with status code: %s",
                name,
                resp.status_code,
            )
            return False
        return True

    def get_dir_id(self, directory: str) -> str:
        """
        Gets the SharePoint Online ID for a given directory to be used for
        subsequent uploads.

        Args:
            directory (str): Name of the directory to get the ID for.

        Returns:
            str: ID for the given directory.
        """
        full_response: dict = self.query_graph(directory)
        if "error" in full_response:
            logger.error("Failure: %s", full_response["error"])
        return full_response.get("id", "")

    def _get_upload_url(self, dir_id: str, file_name: str) -> str:
        """
        Gets the URL to be used for a subsequent file upload.

        Args:
            dir_id (str): ID of the directory to upload to.
            file_name (str): Name of the file being uploaded.

        Returns:
            str: URL to use for the upload session.
        """
        base_url: str = self.endpoint.split("/root:", maxsplit=1)[0]
        endpoint: str = f"{base_url}/items/{dir_id}:/{file_name}:/createUploadSession"
        file_desc: str = "Dashboard script repository archive."
        logger.debug("Getting upload URL from endpoint: %s", endpoint)
        url_payload: dict = {
            "@microsoft.graph.conflictBehavior": "replace",
            "description": file_desc,
            "fileSystemInfo": {"@odata.type": "microsoft.graph.fileSystemInfo"},
            "name": file_name,
        }

        resp: requests.Response = requests.post(
            endpoint,
            headers={
                "Authorization": f"Bearer {self.graph_token['access_token']}",
                "Content-Type": "application/json",
            },
            json=url_payload,
            timeout=DEFAULT_TIMEOUT,
        )

        if not resp.status_code in [200, 201]:
            logger.error(
                "Failed to get an upload URL with status code: %s", resp.status_code
            )
        return resp.json().get("uploadUrl", "")

    def _manage_file_chunks(self, endpoint: str, file_path: str) -> bool:
        """
        Manages the breaking of a file into discrete chunks that SharePoint Online
        will permit to be uploaded and then upload them.

        Args:
            endpoint (str): Endpoint to upload the file to.
            file_path (str): Local path to the file to upload.

        Returns:
            bool: Whether or not the file was successfully uploaded.
        """
        acceptable_codes: list = [200, 201, 202]
        file_stats: os.stat_result = os.stat(file_path)
        file_size: int = file_stats.st_size
        chunks: int = int(file_size / CHUNK_SIZE) + 1
        logger.debug("File will be broken into %s chunks.", chunks)
        with open(file_path, "rb") as file_data:
            start: int = 0
            for chunk_num in range(chunks):
                chunk: bytes = file_data.read(CHUNK_SIZE)
                bytes_read: int = len(chunk)
                upload_range: str = (
                    f"bytes {start}-{start + bytes_read - 1}/{file_size}"
                )
                logger.debug(
                    "Chunk: %s -- Bytes read: %s -- Upload range: %s",
                    chunk_num,
                    bytes_read,
                    upload_range,
                )
                resp: requests.Response = requests.put(
                    endpoint,
                    headers={
                        "Content-Length": str(bytes_read),
                        "Content-Range": upload_range,
                    },
                    data=chunk,
                    timeout=DEFAULT_TIMEOUT,
                )

                if not resp.status_code in acceptable_codes:
                    logger.error(
                        "Failed to upload the chunk with status code: %s",
                        resp.status_code,
                    )
                    return False
                start += bytes_read
        logger.info("Completed the file upload.")
        return True

    def upload_file(self, dir_id: str, file_path: str, file_name: str) -> bool:
        """
        Uploads a file to SharePoint Online.

        Args:
            dir_id (str): ID of the directory to upload to.
            file_path (str): Local path to the file to upload.
            file_name (str): Name of the file being uploaded.

        Returns:
            bool: Whether or not the upload was successful.
        """
        logger.debug(
            "Attempting to upload file %s to directory ID: %s", file_name, dir_id
        )
        upload_url: str = self._get_upload_url(dir_id, file_name)
        logger.info("Using the following upload URL: %s", upload_url)
        if not upload_url:
            return False

        logger.debug("Beginning upload of: %s", file_name)
        return self._manage_file_chunks(upload_url, file_path)

    def _delete_file(self, item_id: str) -> bool:
        """
        Deletes an item with the given ID in SharePoint Online.

        Args:
            item_id (str): ID of the item to delete.

        Returns:
            bool: Whether or not the item was successfully deleted.
        """
        endpoint: str = (
            f'{self.endpoint.split("/root:", maxsplit=1)[0]}/items/{item_id}'
        )
        logger.info("Deleting an old backup with the following URL: %s", endpoint)

        resp: requests.Response = requests.delete(
            endpoint,
            headers={"Authorization": f"Bearer {self.graph_token['access_token']}"},
            timeout=DEFAULT_TIMEOUT,
        )

        if resp.status_code != 204:
            logger.error(
                "Failed to remove the file with status code: %s", resp.status_code
            )
            return False
        return True

    def cleanup_files(self) -> None:
        """
        Cleans up the content in a given SharePoint Online directory.
        """
        logger.info("Doing cleanup on the following directory: %s", self.endpoint)

        dir_content: list = self.query_graph(self.endpoint).get("value", [])
        logger.debug("Found %s items in the directory.", len(dir_content))

        oldest_filename: str = ""
        oldest_id: str = ""
        right_now = datetime.utcnow()

        while len(dir_content) > MAX_BACKUPS:
            oldest_timestamp: datetime = datetime(
                right_now.year + 1, right_now.month, right_now.day, 0, 0, 0
            )
            for single_file in dir_content:
                try:
                    current_file_time: datetime = datetime.strptime(
                        single_file.get("createdDateTime", ""), "%Y-%m-%dT%H:%M:%SZ"
                    )
                    if current_file_time < oldest_timestamp:
                        oldest_timestamp = current_file_time
                        oldest_filename = single_file.get("name", "")
                        oldest_id = single_file.get("id", "")
                except ValueError:
                    logger.error(
                        "Unable to parse the following to a datetime: %s",
                        single_file.get("createdDateTime", ""),
                    )

            logger.info(
                "Oldest file was %s with a timestamp of: %s. It will be deleted.",
                oldest_filename,
                oldest_timestamp,
            )
            if not self._delete_file(oldest_id):
                break

            match: dict = next(
                (file for file in dir_content if file["id"] == oldest_id), {}
            )
            if match:
                dir_content.remove(match)
