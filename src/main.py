#!/usr/bin/env python3
"""
Entrypoint for the GitHub Action.
"""

import logging
import os
import shutil
import sys

from datetime import datetime

from spo_controller import SpoController


ARCHIVE_BASE: str = "/tmp/archive/"

logger: logging.Logger = logging.getLogger()
format_string: str = "%(asctime)s - %(filename)s - %(levelname)s - %(message)s"
log_format: logging.Formatter = logging.Formatter(format_string)
stdout_handler: logging.StreamHandler = logging.StreamHandler(sys.stdout)
stdout_handler.setFormatter(log_format)
logger.addHandler(stdout_handler)
logger.setLevel(logging.INFO)


def manage_spo(tarball_path: str, tarball_name: str) -> None:
    """
    Wrapper for the SharePoint Online connection.

    Args:
        tarball_path (str): Path to the tarball to upload.
        tarball_name (str): Name of the tarball.
    """
    logger.info("Beginning the file upload process.")
    spo: SpoController = SpoController()
    if not spo.validate_info():
        logger.error("Skipping upload since environment information is missing.")
        return
    logger.info("Starting a connection to the Microsoft Graph.")
    spo.connect_graph()
    logger.info("Validating that an access token was obtained.")
    if not spo.check_token():
        logger.error("Failed to retrieve a token for SpO. Moving to cleanup.")
        return

    upload_status: bool = False
    logger.info("Getting the directory ID to target for the upload.")
    logger.debug("Directory is: %s", spo.endpoint.rstrip(":/children/"))
    dir_id: str = spo.get_dir_id(spo.endpoint.rstrip(":/children/"))
    if not dir_id:
        logger.error("Aborting since no directory ID was retrieved.")
        return

    logging.info("Uploading the file in chunks.")
    upload_status = spo.upload_file(dir_id, tarball_path, tarball_name)

    if upload_status:
        logging.info("Beginning the SharePoint Online cleanup process.")
        spo.cleanup_files()
    else:
        logging.info("Ignoring cleanup since the upload was not successful.")

    print("Done with SharePoint Online.")


def remove_file(file_path: str) -> None:
    """
    Deletes the file at the given location.

    Args:
        file_path (str): Path to the file to delete.
    """
    logger.debug("Attempting to remove file: %s", file_path)
    try:
        os.remove(file_path)
        logger.info("File cleaned up successfully: %s", file_path)
    except FileNotFoundError:
        logger.error("Unable to find file to delete: %s", file_path)


def check_debug() -> None:
    """
    Determines if debug logging should be enabled.
    """
    if os.environ.get("DEBUG", "").lower() == "true":
        logger.setLevel(logging.DEBUG)


def validate_dir(dir_path: str) -> None:
    """
    Ensures the given directory exists.

    Args:
        dir_path (str): Path to check.
    """
    if not os.path.exists(dir_path):
        os.makedirs(dir_path)


def main() -> None:
    """
    Entrypoint.
    """
    check_debug()

    file_path: str = os.environ.get("GITHUB_WORKSPACE", "./")
    if not file_path.endswith("/"):
        file_path = f"{file_path}/"
    logger.info("File path is: %s", file_path)

    logger.debug("Validating archive storage directory at: %s", ARCHIVE_BASE)
    validate_dir(ARCHIVE_BASE)

    name_base: str = os.environ.get("ARCHIVE_PREFIX", "repo")
    archive_name: str = f"{name_base}_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}"
    archive_path: str = f"{ARCHIVE_BASE}{archive_name}"

    shutil.make_archive(archive_path, "zip", file_path)
    logger.info("Archive created at: %s", archive_path)

    archive_name = f"{archive_name}.zip"
    archive_path = f"{archive_path}.zip"

    manage_spo(archive_path, archive_name)
    remove_file(archive_path)


if __name__ == "__main__":
    main()
