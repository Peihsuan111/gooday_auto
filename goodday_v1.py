from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service

# from selenium.common.exceptions import TimeoutException
from datetime import date
from dateutil.relativedelta import relativedelta
import datetime
import pandas as pd
import time
import os

# username = str(input("請輸入Goodday帳號： "))
# password = str(input("請輸入Goodday密碼： "))

# today = date.today()
# gates = False
# while gates == False:
#     # input_m = str(input("請輸入月份： "))
#     # input_y = str(input("請輸入年份： "))
#     input_date = input_y + "-" + str(int(input_m)) + "-" + "1"
#     input_date = datetime.datetime.strptime(input_date, "%Y-%m-%d").date()
#     input_date += relativedelta(months=1)
#     first_day = input_date.replace(day=1)
#     last_date = first_day - datetime.timedelta(days=1)
#     first_date = last_date.replace(day=1)
#     last_month = last_date.strftime("%b %#d, %Y")[:3]
#     dif_m = (today.year - last_date.year) * 12 + (today.month - last_date.month)
#     if (dif_m <= 12) & (dif_m >= 0):
#         gates = True
#     else:
#         print("---請重新輸入月份及年份，請勿輸入超過一年內之時間段---")


def init(input_m, input_y):
    username = "sara.huang@datarget.com"
    password = "qazwsx1112"
    input_date = input_y + "-" + str(int(input_m)) + "-" + "1"
    input_date = datetime.datetime.strptime(input_date, "%Y-%m-%d").date()
    input_date += relativedelta(months=1)

    today = date.today()
    first_day = input_date.replace(day=1)  # 9/1
    last_date = first_day - datetime.timedelta(days=1)  # 8/31
    dif_m = (today.year - last_date.year) * 12 + (today.month - last_date.month)
    if not (dif_m <= 12) & (dif_m >= 0):
        print("---請重新輸入月份及年份，請勿輸入超過一年內之時間段---")
    return username, password, dif_m


def Chrome_setting():
    chromeDriverPath = Service(r"chromedriver.exe")

    options = Options()
    options.add_argument("--disable-notifications")
    options.add_argument("--ignore-certificate-errors")
    options.add_argument("--ignore-ssl-errors")
    options.add_argument("--headless=new")  # 打開背景模式
    download = os.getcwd()
    prefs = {"download.default_directory": download}
    options.add_experimental_option("prefs", prefs)
    driver = webdriver.Chrome(options=options, service=chromeDriverPath)
    return driver


def drive_setup(headless=True):
    chrome_options = Options()
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    prefs = {"download.default_directory": os.getcwd()}
    chrome_options.add_experimental_option("prefs", prefs)
    if headless:
        chrome_options.add_argument("--headless")
    path = "chromedriver"  # 'C:\\Users\\USER\\chromedriver.exe'
    driver = webdriver.Chrome(executable_path=path, options=chrome_options)
    return driver


def login(driver, username, password):
    driver.get("https://www.goodday.work/login")
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, "email")))

    # Fill in the login form and submit
    email_input = driver.find_element(By.NAME, "email")
    email_input.send_keys(username)
    password_input = driver.find_element(By.NAME, "password")
    password_input.send_keys(password)
    login_button = driver.find_element(By.CSS_SELECTOR, 'button[type="submit"]')
    login_button.click()

    time.sleep(3)  # to-do: wait till
    return driver


def output_setup(driver):
    driver.get("https://www.goodday.work/analytics/efforts/time-reports-user")

    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located(
            (
                By.XPATH,
                '//*[@id="iv-analytics"]/div/div[1]/div[2]/div[5]',
            )
        )
    )
    generate = driver.find_element(
        By.XPATH,
        "//div[contains(@class,'ui5-button flex m-t-30 s36 primary')]",
    )
    generate.click()

    export = driver.find_element(
        By.XPATH,
        "//div[contains(@class,'ui5-button flex time-report-export-button s26 primary')]",
    )
    export.click()
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located(
            (By.XPATH, "//*[@id='ui-modal']//div[text()='Time Reports']")
        )
    )

    time_reports = driver.find_element(
        By.XPATH, "//*[@id='ui-modal']//div[text()='Time Reports']"
    )
    time_reports.click()
    internal = driver.find_element(
        By.XPATH, "//div[contains(@class,'f-tbl')]//div[text()='Internal cost']"
    )
    internal.click()
    external = driver.find_element(
        By.XPATH, "//div[contains(@class,'f-tbl')]//div[text()='External cost']"
    )
    external.click()
    message = driver.find_element(
        By.XPATH, "//*[@id='ui-modal']//div[text()='Message/Comments']"
    )
    message.click()
    minutes = driver.find_element(
        By.XPATH, "//*[@id='ui-modal']//div[text()='Reported time (Minutes)']"
    )
    minutes.click()
    hours = driver.find_element(
        By.XPATH, "//*[@id='ui-modal']//div[text()='Reported time (Hours)']"
    )
    hours.click()
    return driver


def download_setup(driver, dif_m):
    dates = driver.find_element(By.CLASS_NAME, "gd-daterange-value")
    dates.click()

    left = driver.find_element(
        By.XPATH,
        "//div[contains(@class, 'calendar-content left')]//div[contains(@class, 'arrow ui5-control-button s-28 i5-angle-right')]",
    )
    right = driver.find_element(
        By.XPATH,
        "//div[contains(@class, 'calendar-content right')]//div[contains(@class, 'arrow ui5-control-button s-28 i5-angle-left')]",
    )

    if dif_m <= 3:
        if dif_m == 1:
            left.click()
            left.click()
        elif dif_m == 2:
            left.click()
            right.click()
        elif dif_m == 3:
            right.click()
            right.click()
        elif dif_m == 0:
            left.click()
            left.click()
            left.click()
    elif dif_m > 3:
        for i in range(dif_m - 1):
            right.click()

    # Right calandar first date of the month
    first = driver.find_element(
        By.XPATH, '//div[contains(@class, "calendar-content right")]//div[text()="1"]'
    )
    first.click()
    last = driver.find_element(
        By.XPATH, '//div[contains(@class, "calendar-content left")]//div[text()="1"]'
    )
    last.click()
    return driver


def export(driver):
    export = driver.find_element(
        By.XPATH, '//*[@id="ui-modal"]//div[contains(@class, "gd-btn m-t-10 primary")]'
    )
    export.click()
    time.sleep(3)
    # to-do: close until download complete
    driver.close()


def file_ETL(dif_m):
    curdir = os.getcwd()
    download_month = (date.today() - relativedelta(months=dif_m)).strftime("%b")
    download_yr = (date.today() - relativedelta(months=1)).strftime("%Y")
    old_name = curdir + "/time-reports.csv"
    new_name = curdir + f"/{download_month}-time-reports_raw.csv"
    os.rename(old_name, new_name)
    report = pd.read_csv(new_name)

    report["month"] = report["Report date"].apply(
        lambda x: datetime.datetime.strptime(x, "%Y-%m-%d")
    )
    report["month"] = report["month"].apply(lambda x: x.strftime("%b %#d, %Y")[:3])
    report = report[report["month"] == download_month]
    report[["firm", "project", "drop"]] = report["Project full name"].str.split(
        " > ", n=2, expand=True
    )
    report.drop(["drop", "month"], axis=1, inplace=True)

    report["ID"] = "DTG" + report["User name"].apply(lambda x: x[-3:])
    report["NAME"] = report["User name"].apply(lambda x: x[:-3])

    report["project"] = report.project.str.split("_").str.get(0)
    report.drop(
        ["Report date", "User name", "Task ID", "Task name", "Project full name"],
        axis=1,
        inplace=True,
    )
    report.sort_values("ID", inplace=True)

    report_sum = report.pivot_table(
        index=["ID", "NAME"],
        values="Reported time (Hours)",
        columns="project",
        aggfunc="sum",
        margins=True,
        margins_name="Total",
    ).reset_index()
    report_sum = report_sum[:-1]

    total = report_sum.pop("Total")
    report_sum.insert(2, "Total", total)
    os.remove(new_name)

    # Left Join
    name_list = pd.read_excel("name_list.xlsx", header=0)
    final_report = (
        name_list.merge(report_sum, on="ID", how="left")
        .rename({"NAME_x": "NAME"}, axis=1)
        .drop("NAME_y", axis=1)
    )
    final_report[["Total"]] = final_report[["Total"]].fillna(0)
    final_report.to_excel(
        f"{download_month}{download_yr}_time_reports.xlsx",
        index=False,
        freeze_panes=(1, 3),
        sheet_name=f"{download_month}工時記錄",
        float_format="%.2f",
    )


def runner():
    username, password, dif_m = init("8", "2023")
    driver = Chrome_setting()
    driver = login(driver, username, password)
    driver = output_setup(driver)
    driver = download_setup(driver, dif_m)
    export(driver)
    file_ETL(dif_m)
