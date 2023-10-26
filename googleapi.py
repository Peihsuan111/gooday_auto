#!/usr/bin/python3
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from datetime import date
import datetime
import os, sys

# Create Service
scope = ["https://www.googleapis.com/auth/drive"]
service_account_json_key = "credentials.json"
credentials = service_account.Credentials.from_service_account_file(
    filename=os.path.join(
        os.path.dirname(os.path.realpath(__file__)), service_account_json_key
    ),
    scopes=scope,
)
service = build("drive", "v3", credentials=credentials)


def ListFiles():
    # List Files
    results = (
        service.files()
        .list(
            pageSize=1000,
            fields="nextPageToken, files(id, name, mimeType, size, modifiedTime)",
        )
        .execute()
    )  # , q='name contains "de"'
    items = results.get("files", [])
    print(items)


def UploadFile(filePath, fileName):  # "upload_test.txt"
    # Upload
    print(f"Uploading file {filePath}...")
    uploadFolderId = "1za3m7dnm7Yrn3WujZzVPLhbdvbPfMueH"
    file_metadata = {
        "name": fileName,
        "parents": [uploadFolderId],
    }
    media = MediaFileUpload(
        filePath,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    file = (
        service.files()
        .create(body=file_metadata, media_body=media, fields="id")
        .execute()
    )


if __name__ == "__main__":
    if sys.argv[1] == "l":  # run last month
        thisMonth = datetime.datetime.today().replace(day=1)
        lastMonth = thisMonth - datetime.timedelta(days=1)
        download_year, download_month = lastMonth.year, lastMonth.month
    else:
        download_year, download_month = int(sys.argv[1]), int(sys.argv[2])

    date_naming = date(download_year, download_month, 1).strftime("%b") + str(
        download_year
    )

    dir_path = os.path.dirname(os.path.realpath(__file__))
    file_path = os.path.join(dir_path, f"{date_naming}_time_reports.xlsx")
    file_name = f"{date_naming}_time_reports.xlsx"
    UploadFile(file_path, file_name)
