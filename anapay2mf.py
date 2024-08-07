import os
import time
import logging
import traceback
from datetime import datetime, timedelta
from typing import List, Optional

import imaplib
import email
from email.header import decode_header

import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.errors import HttpError

from dateutil import parser
from dotenv import load_dotenv

from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import helium

from dataclasses import dataclass


# 環境変数を読み込む
load_dotenv()

# 環境変数から情報を取得
SHEET_ID = os.getenv("SHEET_ID")
EMAIL = os.getenv("EMAIL") # gmail
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD") # gmail
EMAIL_MF = os.getenv("EMAILMF") # Moneyforward
PASSWORD = os.getenv("PASSWORD")
GOOGLE_APPLICATION_CREDENTIALS = os.getenv("GOOGLE_APPLICATION_CREDENTIALS")
SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
MAILBOX = os.getenv("GMAIL_MAILBOXNAME")

# 必須環境変数のチェック
required_env_vars = ['SHEET_ID', 'EMAIL', 'EMAIL_PASSWORD', 'GOOGLE_APPLICATION_CREDENTIALS']
missing_vars = [var for var in required_env_vars if not os.getenv(var)]
if missing_vars:
    raise EnvironmentError(f"Required environment variables are missing: {', '.join(missing_vars)}")


@dataclass
class ANAPay:
    """ANA Pay information"""
    email_date: datetime = None
    date_of_use: datetime = None
    amount: int = 0
    store: str = ""
    email_id: str = ""

    def values(self) -> tuple[str, str, str, str]:
        """return tuple of values for spreadsheet"""
        return self.email_date_str, self.date_of_use_str, self.amount, self.store

    @property
    def email_date_str(self) -> str:
        return f"{self.email_date:%Y-%m-%d %H:%M:%S}"

    @property
    def date_of_use_str(self) -> str:
        return f"{self.date_of_use:%Y-%m-%d %H:%M:%S}"


def get_mail_info(msg, email_id) -> Optional[ANAPay]:
    """
    1件のメールからANA Payの利用情報を取得して返す
    """
    ana_pay = ANAPay(email_id=email_id)
    for header in msg["headers"]:
        if header["name"] == "Date":
            date_str = header["value"].replace(" +0900 (JST)", "")
            ana_pay.email_date = parser.parse(date_str)

    # 本文から日時、金額、店舗を取り出す
    body = msg["body"]
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


def get_anapay_info(imap_server, username, password, after: str) -> List[ANAPay]:
    """
    IMAPを使用してGmailからANA Payの利用履歴を取得する
    """
    ana_pay_list = []
    mail = imaplib.IMAP4_SSL(imap_server)

    try:
        mail.login(username, password)
        mail.select(MAILBOX)

    except Exception as e:
        logging.error(f"IMAP login/select exception: {e}")
        return ana_pay_list

    # 日付をIMAPの検索形式に変換
    since_date = datetime.strptime(after, "%d-%b-%Y")
    since_str = since_date.strftime("%d-%b-%Y")

    # 検索条件を設定（件名に「[ANA Pay]」を含む）
    query = f'(FROM "payinfo@121.ana.co.jp" SUBJECT "[ANA Pay]" SINCE {since_str})'
    logging.info(f"IMAP search query: {query}")

    try:
        result, data = mail.search(None,query)
        logging.info(f"IMAP search result: {result}, data: {data}")
    except Exception as e:
        logging.error(f"IMAP search exception: {e}")
        return ana_pay_list

    if result != 'OK':
        logging.error(f"IMAP search failed with result: {result}, data: {data}")
        return ana_pay_list

    email_ids = data[0].split()

    for email_id in reversed(email_ids):
        result, msg_data = mail.fetch(email_id, "(RFC822)")
        msg = email.message_from_bytes(msg_data[0][1])

        # 件名をデコードして確認
        subject, encoding = decode_header(msg['Subject'])[0]
        if isinstance(subject, bytes):
            subject = subject.decode(encoding if encoding else 'utf-8')
        if "［ANA Pay］ご利用のお知らせ" not in subject:
            continue

        # 本文をデコードして「ご利用日時」を含むか確認
        if msg.is_multipart():
            for part in msg.walk():
                if part.get_content_type() == "text/plain":
                    body = part.get_payload(decode=True).decode(part.get_content_charset())
                    if "ご利用日時" in body:
                        email_data = {
                            "headers": [{"name": k, "value": v} for k, v in msg.items()],
                            "body": body,
                        }
                        ana_pay = get_mail_info(email_data, email_id)
                        if ana_pay:
                            ana_pay_list.append(ana_pay)
                        break
        else:
            body = msg.get_payload(decode=True).decode(msg.get_content_charset())
            if "ご利用日時" in body:
                email_data = {
                    "headers": [{"name": k, "value": v} for k, v in msg.items()],
                    "body": body,
                }
                ana_pay = get_mail_info(email_data, email_id)
                if ana_pay:
                    ana_pay_list.append(ana_pay)

    mail.close()
    mail.logout()
    return ana_pay_list





def get_last_email_date(records: list[dict[str, str]]):
    """get last email date for gmail search"""
    after = "20-Mar-2024"
    if records:
        last_email_date = parser.parse(records[-1]["email_date"])
        after = f"{last_email_date.strftime('%d-%b-%Y')}"
    return after


def gmail2spredsheet(worksheet):
    """IMAPからANA Payの利用履歴を取得しスプレッドシートに書き込む"""
    # get all records from spreadsheet
    records = worksheet.get_all_records()
    logging.info("Records in spreadsheet: %d", len(records))

    # get last day from records
    after = get_last_email_date(records)
    logging.info("Last day on spreadsheet: %s", after)
    email_date_set = set(parser.parse(r["email_date"]) for r in records)

    # get ANA Pay email from IMAP
    ana_pay_list = get_anapay_info("imap.gmail.com", EMAIL, EMAIL_PASSWORD, after)
    logging.info("ANA Pay emails: %d", len(ana_pay_list))

    # add ANA Pay record to spreadsheet
    count = 0
    for ana_pay in ana_pay_list:
        # メールの日付が存在しない場合はレコードを追加
        if ana_pay.email_date not in email_date_set:
            try:
                worksheet.append_row(ana_pay.values(), value_input_option="USER_ENTERED")
                count += 1
                logging.info("Record added to spreadsheet: %s", ana_pay.values())
                time.sleep(1)
            except Exception as e:
                logging.error(f"Error adding record to spreadsheet: {e}")
    logging.info("Records added to spreadsheet: %d", count)

    # 既読にするメールのIDをリストアップ
    email_ids_to_mark_read = [ana_pay.email_id for ana_pay in ana_pay_list if ana_pay.email_date not in email_date_set]

    # メールを既読にする
    for email_id in email_ids_to_mark_read:
        try:
            mark_as_read("imap.gmail.com", EMAIL, EMAIL_PASSWORD, email_id)
            logging.info("Email marked as read: %s", email_id)
        except Exception as e:
            logging.error(f"Error marking email as read: {e}")


def mark_as_read(imap_server, username, password, email_id):
    """
    指定したメールを既読にする
    """
    mail = imaplib.IMAP4_SSL(imap_server)
    try:
        mail.login(username, password)
        mail.select("inbox")
        mail.store(email_id, '+FLAGS', '\\Seen')
    except Exception as e:
        logging.error(f"IMAP mark as read exception: {e}")
    finally:
        mail.close()
        mail.logout()


def save_screenshot(driver, filename):
    path = os.path.join("/app/screenshots", filename)
    os.makedirs(os.path.dirname(path), exist_ok=True)
    driver.save_screenshot(path)
    logging.info(f"スクリーンショットを保存しました: {path}")


def login_mf():
    """login moneyforward sbi"""


    if not EMAIL_MF or not PASSWORD:
        logging.error("MoneyforwadのログインEMAILまたはPASSWORDが設定されていません。")
        return

    logging.info(f"使用するEMAIL: {EMAIL_MF}")
    logging.info(f"使用するPASSWORD: {'*' * len(PASSWORD)}")

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
    options.add_argument(
        "--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3")
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
        email_input.send_keys(EMAIL_MF)
        logging.info(f"メールアドレスを入力: {EMAIL_MF}")
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
        password_input.send_keys(PASSWORD)
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
            password_input.send_keys(PASSWORD)
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
                        password_input.send_keys(PASSWORD)
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
                    account_button = driver.find_element(By.XPATH,
                                                         "/html/body/main/div/div/div[2]/div/section/div/div/form/button")
                    account_button.click()
                    WebDriverWait(driver, 30).until(
                        EC.presence_of_element_located((By.NAME, "mfid_user[password]"))
                    )
                    password_input = driver.find_element(By.NAME, "mfid_user[password]")
                    password_input.send_keys(PASSWORD)
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


def add_mf_record(dt: datetime, amount: int, store: str, store_info: Optional[dict]):
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

        # ANA Payの選択
        payment_select = driver.find_element(By.ID, "user_asset_act_sub_account_id_hash")
        select = Select(payment_select)
        for option in select.options:
            if option.text.startswith("ANA Pay"):
                select.select_by_visible_text(option.text)
                break

        if store_info:
            # カテゴリー選択
            l_category = driver.find_element(By.CSS_SELECTOR, "#js-large-category-selected")
            l_category.click()
            save_screenshot(driver, "cliked_large_category.png")

            l_category_option = driver.find_element(By.XPATH,
                                                    f"//a[@class='l_c_name' and text()='{store_info['大項目']}']")
            l_category_option.click()
            save_screenshot(driver, "selected_large_category.png")

            m_category = driver.find_element(By.CSS_SELECTOR, "#js-middle-category-selected")
            m_category.click()
            m_category_option = driver.find_element(By.XPATH,
                                                    f"//a[@class='m_c_name' and text()='{store_info['中項目']}']")
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


def main():
    try:
        # ログ設定
        logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(name)s - %(levelname)s - %(message)s")

        # gspread クライアントの初期化
        creds = Credentials.from_service_account_file(GOOGLE_APPLICATION_CREDENTIALS, scopes=SCOPES)
        gc = gspread.authorize(creds)

        sheet = gc.open_by_key(SHEET_ID)
        anapay_sheet = sheet.worksheet("ANAPay")
        store_sheet = sheet.worksheet("ANAPayStore")
        store_dict = {store["store"]: store for store in store_sheet.get_all_records()}

        # データの処理
        gmail2spredsheet(anapay_sheet)
        spreadsheet2mf(anapay_sheet, store_dict)

    except gspread.exceptions.SpreadsheetNotFound as e:
        logging.error(f'Spreadsheet not found: {e}')
    except HttpError as error:
        logging.error(f'An error occurred with Google Sheets API: {error}')
    except Exception as e:
        logging.error(f'An unexpected error occurred: {e}')
        traceback.print_exc()  # 詳細なスタックトレースを表示する


if __name__ == "__main__":
    main()
