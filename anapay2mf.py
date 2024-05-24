import base64
import logging
import os
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
import time

import gspread
import helium
from dateutil import parser
from dotenv import load_dotenv
from google.auth.transport.requests import Request
from google.auth.exceptions import RefreshError
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError


from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException

import quickstart

# .env ファイルをロード
load_dotenv()

# 環境変数から SHEET_ID を取得
SHEET_ID = os.getenv("SHEET_ID")
SHEET_NAME = "ANAPay"

SCOPES = [
    "https://www.googleapis.com/auth/gmail.readonly",
    "https://www.googleapis.com/auth/spreadsheets",
]

MF_URL = "https://ssnb.x.moneyforward.com/cf"

TOKEN_JSON = "/app/token.json"
CREDENTIALS_JSON = "/app/credentials.json"

format = "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
logging.basicConfig(format=format, level=logging.INFO)

creds = None

@dataclass
class ANAPay:
    """ANA Pay information"""

    email_date: datetime = None
    date_of_use: datetime = None
    amount: int = 0
    store: str = ""

    def values(self) -> tuple[str, str, str, str]:
        """return tuple of values for spreadsheet"""
        return self.email_date_str, self.date_of_use_str, self.amount, self.store

    @property
    def email_date_str(self) -> str:
        return f"{self.email_date:%Y-%m-%d %H:%M:%S}"

    @property
    def date_of_use_str(self) -> str:
        return f"{self.date_of_use:%Y-%m-%d %H:%M:%S}"


def get_mail_info(res: dict) -> ANAPay | None:
    """
    1件のメールからANA Payの利用情報を取得して返す
    """
    ana_pay = ANAPay()
    for header in res["payload"]["headers"]:
        if header["name"] == "Date":
            date_str = header["value"].replace(" +0900 (JST)", "")
            ana_pay.email_date = parser.parse(date_str)

    # 本文から日時、金額、店舗を取り出す
    # ご利用日時：2023-06-28 22:46:19
    # ご利用金額：44,308円
    # ご利用店舗：SMOKEBEERFACTORY OTSUKATE
    data = res["payload"]["body"]["data"]
    body = base64.urlsafe_b64decode(data).decode()
    for line in body.splitlines():
        if line.startswith("ご利用"):
            key, value = line.split("：")
            if key == "ご利用日時":
                ana_pay.date_of_use = parser.parse(value)
            elif key == "ご利用金額":
                ana_pay.amount = int(value.replace(",", "").replace("円", ""))
            elif key == "ご利用店舗":
                ana_pay.store = value
    return ana_pay


def get_anapay_info(after: str) -> list[ANAPay]:
    """
    gmailからANA Payの利用履歴を取得する
    """
    ana_pay_list = []

    creds = Credentials.from_authorized_user_file(TOKEN_JSON, SCOPES)

    # Gmail APIサービスオブジェクトの作成
    service = build('gmail', 'v1', credentials=creds)
    # https://developers.google.com/gmail/api/reference/rest/v1/users.messages/list
    query = f"from:payinfo@121.ana.co.jp subject:ご利用のお知らせ after:{after}"
    results = service.users().messages().list(userId="me", q=query).execute()
    messages = results.get("messages", [])
    for message in reversed(messages):
        # https://developers.google.com/gmail/api/reference/rest/v1/users.messages/get
        res = service.users().messages().get(userId="me", id=message["id"]).execute()
        ana_pay = get_mail_info(res)
        if ana_pay:
            ana_pay_list.append(ana_pay)
    return ana_pay_list

    after = "2024/03/20"


def get_last_email_date(records: list[dict[str, str]]):
    """get last email date for gmail search"""
    after = "2024/03/20"
    if records:
        last_email_date = parser.parse(records[-1]["email_date"])
        after = f"{last_email_date:%Y/%m/%d}"
    return after


def gmail2spredsheet(worksheet):
    """gmailからANA Payの利用履歴を取得しスプレッドシートに書き込む"""
    # get all records from spreadsheet
    records = worksheet.get_all_records()
    logging.info("Records in spreadsheet: %d", len(records))

    # get last day from records
    after = get_last_email_date(records)
    logging.info("Last day on spreadsheet: %s", after)
    email_date_set = set(parser.parse(r["email_date"]) for r in records)

    # get ANA Pay email from Gamil
    ana_pay_list = get_anapay_info(after)
    logging.info("ANA Pay emails: %d", len(ana_pay_list))

    # add ANA Pay record to spreadsheet
    count = 0
    for ana_pay in ana_pay_list:
        # メールの日付が存在しない場合はレコードを追加
        if ana_pay.email_date not in email_date_set:
            worksheet.append_row(ana_pay.values(), value_input_option="USER_ENTERED")
            count += 1
            logging.info("Record added to spreadsheet: %s", ana_pay.values())
            time.sleep(1)
    logging.info("Records added to spreadsheet: %d", count)


def save_screenshot(driver, filename):
    path = os.path.join("/app/screenshots", filename)
    os.makedirs(os.path.dirname(path), exist_ok=True)
    driver.save_screenshot(path)
    logging.info(f"スクリーンショットを保存しました: {path}")


def login_mf():
    """login moneyforward sbi"""

    email = os.getenv("EMAIL")
    password = os.getenv("PASSWORD")

    if not email or not password:
        logging.error("EMAILまたはPASSWORDが設定されていません。")
        return

    logging.info(f"使用するEMAIL: {email}")
    logging.info(f"使用するPASSWORD: {'*' * len(password)}")

    # SeleniumでChromiumを使用する設定
    options = Options()
    options.headless = True
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--remote-debugging-port=9222")
    options.add_argument("--disable-extensions")
    options.add_argument("--disable-software-rasterizer")
    options.add_argument("--headless")
    options.add_argument("--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3")
    options.add_argument("--lang=ja-JP")
    service = ChromeService(executable_path='/usr/bin/chromedriver')

    driver = webdriver.Chrome(service=service, options=options)

    logging.info("Login to moneyfoward")
    driver.get("https://id.moneyforward.com/sign_in")

    try:
        logging.info("ログインID入力ページを待機中")
        # ログインIDを入力してログインボタンをクリック
        email_input = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.NAME, "mfid_user[email]"))
        )
        email_input.send_keys(email)
        logging.info(f"メールアドレスを入力: {email}")
        login_button = driver.find_element(By.ID, "submitto")
        login_button.click()

        logging.info("パスワード入力ページを待機中")
        # パスワード入力ページが読み込まれるのを待つ
        save_screenshot(driver, "before_password.png")
        try:
            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.NAME, "mfid_user[password]"))
            )
        except TimeoutException:
            logging.error("パスワード入力ページの読み込みがタイムアウトしました")
            save_screenshot(driver, "password_page_timeout.png")
            return

        logging.info("パスワード入力中")
        # パスワードを入力してログインボタンをクリック
        password_input = driver.find_element(By.NAME, "mfid_user[password]")
        password_input.send_keys(password)
        logging.info("パスワードを入力")
        login_button = driver.find_element(By.ID, "submitto")
        login_button.click()

        logging.info("ログイン後のページを待機中")
        try:
            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'container-large')]"))
            )
            save_screenshot(driver, "after_login.png")
        except TimeoutException:
            logging.info("ページの読み込みが完了しませんでしたが、処理を続行します")
            save_screenshot(driver, "after_login_timeout.png")

        # 指定されたURLに遷移
        logging.info("指定されたURLに遷移中")
        driver.get("https://moneyforward.com/cf")

        logging.info("パスワード入力中")
        # パスワードを入力してログインボタンをクリック
        try:
            password_input = WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.NAME, "mfid_user[password]"))
            )
            password_input.send_keys(password)
            logging.info("パスワードを再入力")
            login_button = driver.find_element(By.ID, "submitto")
            login_button.click()
        except TimeoutException:
            logging.error("パスワード入力要素が見つかりませんでした")
            save_screenshot(driver, "password_input_not_found.png")
            # アカウント選択画面が表示されているか確認
            try:
                account_selection = driver.find_element(By.XPATH, "/html/body/main/div/div/div[2]/div/section/h1")
                if account_selection and account_selection.text == "アカウントを選択する":
                    logging.info("アカウント選択画面が表示されました")
                    save_screenshot(driver, "account_selection.png")
                    account_button = driver.find_element(By.XPATH,
                                                         "/html/body/main/div/div/div[2]/div/section/div/div/form/button")
                    account_button.click()
                    try:
                        WebDriverWait(driver, 30).until(
                            EC.presence_of_element_located((By.NAME, "mfid_user[password]"))
                        )
                        password_input = driver.find_element(By.NAME, "mfid_user[password]")
                        password_input.send_keys(password)
                        logging.info("パスワードを再入力")
                        login_button = driver.find_element(By.ID, "submitto")
                        login_button.click()
                        WebDriverWait(driver, 30).until(
                            EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'container-large')]"))
                        )
                        save_screenshot(driver, "after_account_selection.png")
                    except TimeoutException:
                        logging.info("アカウント選択後のパスワード入力要素が見つかりませんでしたが、処理を続行します")
                        save_screenshot(driver, "account_password_input_not_found.png")
                        driver.get("https://moneyforward.com/cf")
            except NoSuchElementException:
                logging.error("アカウント選択画面も表示されていませんでした")
                return

        logging.info("「手入力」ボタンを待機中")
        try:
            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.XPATH, "//*[@id='kakeibo']/section/div[1]/div[1]/div/button"))
            )
            save_screenshot(driver, "before_button_click.png")

            logging.info("「手入力」ボタンをクリック")
            button = driver.find_element(By.XPATH, "//*[@id='kakeibo']/section/div[1]/div[1]/div/button")
            button.click()
        except TimeoutException:
            logging.info("「手入力」ボタンが見つかりませんでしたが、処理を続行します")
            save_screenshot(driver, "button_not_found.png")
        except NoSuchElementException:
            logging.info("「手入力」ボタンが見つかりませんでしたが、処理を続行します")
            save_screenshot(driver, "button_not_found.png")
            try:
                account_selection = driver.find_element(By.XPATH, "/html/body/main/div/div/div[2]/div/section/h1")
                if account_selection and account_selection.text == "アカウントを選択する":
                    logging.info("アカウント選択画面が表示されました")
                    save_screenshot(driver, "account_selection.png")
                    account_button = driver.find_element(By.XPATH, "/html/body/main/div/div/div[2]/div/section/div/div/form/button")
                    account_button.click()
                    WebDriverWait(driver, 30).until(
                        EC.presence_of_element_located((By.NAME, "mfid_user[password]"))
                    )
                    password_input = driver.find_element(By.NAME, "mfid_user[password]")
                    password_input.send_keys(password)
                    logging.info("パスワードを再入力")
                    login_button = driver.find_element(By.ID, "submitto")
                    login_button.click()
                    WebDriverWait(driver, 30).until(
                        EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'container-large')]"))
                    )
                    save_screenshot(driver, "after_account_selection.png")
            except Exception as e:
                logging.error(f"アカウント選択後の操作に失敗しました: {e}")
                save_screenshot(driver, "account_selection_error.png")
        except Exception as e:
            logging.error(f"手入力ボタンのクリックに失敗しました: {e}")
            save_screenshot(driver, "button_click_error.png")

        # トップ画面に遷移した場合の処理
        try:
            top_page_indicator = driver.find_element(By.XPATH, "//*[@id='cf-manual-entry']/h2")
            if (top_page_indicator and top_page_indicator.text == "カンタン入力"):
                logging.info("トップ画面が表示されました。再度指定されたURLに遷移します")
                save_screenshot(driver, "top_page_detected.png")
                driver.get("https://moneyforward.com/cf")
                WebDriverWait(driver, 30).until(
                    EC.presence_of_element_located((By.XPATH, "//*[@id='kakeibo']/section/div[1]/div[1]/div/button"))
                )
                save_screenshot(driver, "before_button_click_after_redirect.png")
                button = driver.find_element(By.XPATH, "//*[@id='kakeibo']/section/div[1]/div[1]/div/button")
                button.click()
        except NoSuchElementException:
            logging.info("トップ画面ではありません。処理を続行します")
        except TimeoutException:
            logging.info("指定されたURLに遷移後、ボタンが見つかりませんでしたが、処理を続行します")
            save_screenshot(driver, "button_not_found_after_redirect.png")

    except TimeoutException as e:
        logging.error(f"ログインプロセスがタイムアウトしました: {e}")
        save_screenshot(driver, "timeout_error.png")
        return

    helium.set_driver(driver)


def add_mf_record(dt: datetime, amount: int, store: str, store_info: dict | None):
    """
    add record to moneyfoward
    """
    try:
        driver = helium.get_driver()
        # 「手入力」ボタンをクリック
        helium.click("手入力")
        logging.info(f"手入力をクリック")
        save_screenshot(helium.get_driver(), "clicked_manual_input.png")

        # 日付を入力
        date_input = driver.find_element(By.NAME, "user_asset_act[updated_at]")
        date_input.clear()
        date_input.send_keys(f"{dt:%Y/%m/%d}")
        logging.info("日付を入力")
        save_screenshot(driver, "added_date_input.png")

        # カレンダーポップアップを閉じるために指定された要素をクリック
        popup_closer = driver.find_element(By.XPATH, "//*[@id=\"important\"]/label")
        popup_closer.click()
        logging.info(f"カレンダーポップアップを閉じるために指定された要素をクリック")
        save_screenshot(driver, "closed_calendar_popup.png")

        # 支出金額を入力
        helium.write(amount, into="支出金額")
        logging.info(f"支出金額を入力")
        save_screenshot(helium.get_driver(), "added_expense_amount_input.png")

        if store_info:
            # カテゴリー選択
            l_category = driver.find_element(By.CSS_SELECTOR, "#js-large-category-selected")
            l_category.click()
            save_screenshot(driver, "cliked_large_category.png")

            l_category_option = driver.find_element(By.XPATH, f"//a[@class='l_c_name' and text()='{store_info['大項目']}']")
            l_category_option.click()
            save_screenshot(driver, "selected_large_category.png")

            m_category = driver.find_element(By.CSS_SELECTOR, "#js-middle-category-selected")
            m_category.click()
            m_category_option = driver.find_element(By.XPATH, f"//a[@class='m_c_name' and text()='{store_info['中項目']}']")
            m_category_option.click()

            # 店名を入力
            store_name = store_info.get("店名") or store
            content_input = driver.find_element(By.NAME, "user_asset_act[content]")
            content_input.clear()
            content_input.send_keys(store_name)
        else:
            content_input = driver.find_element(By.NAME, "user_asset_act[content]")
            content_input.clear()
            content_input.send_keys(store)

        # 保存ボタンをクリック
        helium.click("保存する")
        logging.info(f"Record added to moneyforward: {dt:%Y/%m/%d}, {amount}, {store}")

        # 「続けて入力する」ボタンを待機してクリック
        helium.wait_until(helium.Button("続けて入力する").exists)
        helium.click("続けて入力する")
        return True
    except Exception as e:
        logging.error(f"Error adding record to moneyforward: {e}")
        save_screenshot(helium.get_driver(), "add_record_error.png")
        return False

def spreadsheet2mf(worksheet, store_dict: dict[str, dict[str, str]]) -> None:
    """スプレッドシートからmoneyfowardに書き込む"""

    records = worksheet.get_all_records()

    # すべてmoneyforwardに登録済みならなにもしない
    if all(record["mf"] == "done" for record in records):
        logging.error(f"Done. all records are finished")
        return

    login_mf()  # login to moneyfoward
    added = 0
    for count, record in enumerate(records):
        if record["mf"] != "done":
            date_of_use = parser.parse(record["date_of_use"])
            amount = int(record["amount"])
            store = record["store"]
            success = add_mf_record(date_of_use, amount, store, store_dict.get(store))
            logging.info(f"add_mf_record returned: {success}")
            if success:
                try:
                    # update spread sheets for "done" message
                    row = count + 2  # Adjust for 0-based index and header row
                    logging.info(f"Updating cell for record {row}")
                    worksheet.update_cell(row, 5, "done")
                    added += 1
                    time.sleep(1)  # Wait for 1 second to avoid rate limiting
                except Exception as e:
                    logging.error(f"Error updating cell for record {row}: {e}")
    helium.kill_browser()

    logging.info(f"Records added to moneyforward: {added}")


def get_credentials():
    """サービスアカウントの認証情報を返す"""
    creds = Credentials.from_authorized_user_file(TOKEN_JSON, SCOPES)
    return creds


def main():

    if os.path.exists(TOKEN_JSON):
        creds = Credentials.from_authorized_user_file(TOKEN_JSON, SCOPES)

    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                CREDENTIALS_JSON, SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open(TOKEN_JSON, 'w') as token:
            token.write(creds.to_json())

    # Gmail APIサービスオブジェクトの作成
    service = build('gmail', 'v1', credentials=creds)

    try:
        creds = Credentials.from_authorized_user_file(TOKEN_JSON, SCOPES)
        service = build('gmail', 'v1', credentials=creds)
        results = service.users().labels().list(userId='me').execute()
    except RefreshError:
        # recreate token
        Path(TOKEN_JSON).unlink(missing_ok=True)
        quickstart.main()

    try:
        # Google Sheets APIへのアクセス
        gc = gspread.oauth(
            credentials_filename=CREDENTIALS_JSON, authorized_user_filename=TOKEN_JSON
        )

        sheet = gc.open_by_key(SHEET_ID)
        anapay_sheet = sheet.worksheet("ANAPay")
        store_sheet = sheet.worksheet("ANAPayStore")
        store_dict = {store["store"]: store for store in store_sheet.get_all_records()}

        # データの処理
        gmail2spredsheet(anapay_sheet)
        spreadsheet2mf(anapay_sheet, store_dict)

    except gspread.exceptions.SpreadsheetNotFound as e:
        print(f'Spreadsheet not found: {e}')
    except HttpError as error:
        print(f'An error occurred with Google Sheets API: {error}')
    except Exception as e:
        print(f'An unexpected error occurred: {e}')
if __name__ == "__main__":
    main()
