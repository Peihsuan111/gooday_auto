"""_summary_

Example:
    # Run last month
    >> python3 goodday_api.py l
    
    # Run specific month
    >> python3 goodday_api.py 2023 8

"""
import requests
import os
import pandas as pd
from datetime import date
import datetime
import sys

if sys.argv[1] == "l":  # run last month
    thisMonth = datetime.datetime.today().replace(day=1)
    lastMonth = thisMonth - datetime.timedelta(days=1)
    download_year, download_month = lastMonth.year, lastMonth.month
else:
    download_year, download_month = int(sys.argv[1]), int(sys.argv[2])

start_date = datetime.date(download_year, download_month, 1).strftime("%Y-%m-%d")
end_date = datetime.date(download_year, download_month + 1, 1).strftime("%Y-%m-%d")
print(f"Running {download_year}, {download_month}\nFrom {start_date} to {end_date}")
gooday_token = "8d73cf48e9d04fb3925744c8d2381fab"


def get_reports(start_date, end_date):
    # reports
    # start_date = "2023-09-01"
    # end_date = "2023-10-01"
    print("get reports...")
    url = f"https://api.goodday.work/2.0//time-reports?gd-api-token={gooday_token}&startDate={start_date}&endDate={end_date}"
    payload = {}
    headers = {"x-token": "fake-super-secret-token"}
    response = requests.request("POST", url, headers=headers, data=payload)
    data = response.json()
    return data


def get_projects():
    # projects
    print("get projects...")
    url = f"https://api.goodday.work/2.0/projects?gd-api-token={gooday_token}"
    payload = {}
    headers = {}
    response = requests.request("GET", url, headers=headers, data=payload)
    res = response.json()
    projects = dict()
    for i in res:
        projects[i["id"]] = {"parentProjectId": i["parentProjectId"], "name": i["name"]}
    return projects


def get_users():
    # users
    print("get users...")
    url = f"https://api.goodday.work/2.0/users?gd-api-token={gooday_token}"
    payload = {}
    headers = {}
    response = requests.request("GET", url, headers=headers, data=payload)
    res = response.json()
    users = dict()
    for i in res:
        users[i["id"]] = i["name"]
    return users


def lastProjectId(id):
    if id not in projects:
        queryProjectId(id)
    return projects[id]["parentProjectId"]


def queryProjectId(projectId):
    print(f"/////query new project Id {projectId}/////")
    url = (
        f"https://api.goodday.work/2.0//project/{projectId}?gd-api-token={gooday_token}"
    )
    payload = {}
    headers = {}
    response = requests.request("GET", url, headers=headers, data=payload)
    res = response.json()
    projects[res["id"]] = {
        "parentProjectId": res["parentProjectId"],
        "name": res["name"],
    }
    return res


def findProjectFullName(id):
    projNm = projects[id]["name"] if id in projects else queryProjectId(id)["name"]
    all_proj = [projNm]
    while lastProjectId(id) is not None:
        id = lastProjectId(id)
        projNm = projects[id]["name"]
        all_proj.append(projNm)
    all_proj.reverse()
    return " > ".join(all_proj)


def clean_data(df):
    print("Cleanning data...")
    df["month"] = df["Report date"].apply(
        lambda x: datetime.datetime.strptime(x, "%Y-%m-%d")
    )
    df["month"] = df["month"].apply(lambda x: x.strftime("%b %#d, %Y")[:3])
    df[["firm", "project", "drop"]] = df["Project full name"].str.split(
        " > ", n=2, expand=True
    )
    df.drop(["drop", "month"], axis=1, inplace=True)

    df["ID"] = "DTG" + df["User name"].apply(lambda x: x[-3:])
    df["NAME"] = df["User name"].apply(lambda x: x[:-3])

    df["project"] = df.project.str.split("_").str.get(0)
    df.drop(
        ["Report date", "User name", "Project full name"],
        axis=1,
        inplace=True,
    )
    df.sort_values("ID", inplace=True)

    df_sum = df.pivot_table(
        index=["ID", "NAME"],
        values="Reported time (Hours)",
        columns="project",
        aggfunc="sum",
        margins=True,
        margins_name="Total",
    ).reset_index()
    df_sum = df_sum[:-1]

    curdir = os.getcwd()
    date_naming = date(download_year, download_month, 1).strftime("%b") + str(
        download_year
    )
    file_path = os.path.join(curdir, f"{date_naming}_time_reports.xlsx")
    print(f"Output to {file_path}...")
    total = df_sum.pop("Total")
    df_sum.insert(2, "Total", total)

    df_sum.to_excel(
        file_path,
        index=False,
        freeze_panes=(1, 3),
        sheet_name=f"{date_naming}工時記錄",
        float_format="%.2f",
    )


data = get_reports(start_date, end_date)
users = get_users()
projects = get_projects()
reports = list()
for i in data:
    obj = dict()
    obj["Report date"] = i["reportDate"]
    obj["User name"] = users[i["userId"]] if i["userId"] in users else ""
    obj["Reported time (Hours)"] = i["reported"] / 60
    obj["Project full name"] = findProjectFullName(i["projectId"])
    reports.append(obj)

df = pd.DataFrame(reports)  # .to_csv("time-reports.csv")
clean_data(df)
