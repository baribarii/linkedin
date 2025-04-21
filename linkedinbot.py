#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
LinkedIn 포스트 데이터를 주기적으로 Google 스프레드시트에 기록
로컬(macOS): 키체인 ‘LinkedIn’ 항목(email / password) 사용
CI(GitHub Actions): 환경변수 + Secrets(Base64) 사용
"""

import os
import sys
import time
import datetime
import glob
import re
import base64
import platform
import pandas as pd

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

from googleapiclient.discovery import build
from google.oauth2.service_account import Credentials

# macOS 키체인
try:
    import keyring
except ImportError:
    keyring = None  # CI 환경에서는 keyring 모듈이 없어도 됨

# ------------------------------------------------
# 1. Google Service Account 인증
# ------------------------------------------------
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

LOCAL_KEY_PATH = os.path.expanduser("~/Downloads/my_new_key.json")
CI_KEY_PATH    = "service_account_temp.json"

if os.path.exists(LOCAL_KEY_PATH):
    SERVICE_ACCOUNT_FILE = LOCAL_KEY_PATH
else:
    b64_json = os.getenv("LINKEDIN_GOOGLESHEET_API")
    if not b64_json:
        raise EnvironmentError("LINKEDIN_GOOGLESHEET_API 환경변수가 없습니다.")
    decoded = base64.b64decode(b64_json) if not b64_json.strip().startswith("{") else b64_json.encode()
    with open(CI_KEY_PATH, "wb") as f:
        f.write(decoded)
    SERVICE_ACCOUNT_FILE = CI_KEY_PATH

creds   = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
service = build('sheets', 'v4', credentials=creds)

SPREADSHEET_ID = '1fQTqTrNGwSNGi9EzyK8A2ZqU48IbXG-YrL2ImhXm74w'
SHEET_NAME     = '시트4'

# ------------------------------------------------
# 2. LinkedIn 로그인 정보
# ------------------------------------------------
def get_linkedin_credentials():
    email    = os.getenv("LINKEDIN_EMAIL")
    password = os.getenv("LINKEDIN_PASSWORD")

    if platform.system() == "Darwin" and keyring:
        try:
            if not email:
                email = keyring.get_password("LinkedIn", "email")
            if not password:
                password = keyring.get_password("LinkedIn", "password")
        except Exception as e:
            print("[WARN] 키체인 읽기 실패:", e)
    return email, password

# ------------------------------------------------
# 3. 스프레드시트 유틸
# ------------------------------------------------
def get_analytics_url() -> str | None:
    cell = f"{SHEET_NAME}!C2"
    values = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID, range=cell
    ).execute().get("values", [])
    if not values:
        print("C2 셀에 URL이 없습니다.")
        return None

    feed_url = values[0][0].strip()
    if "urn:li:activity:" in feed_url:
        act_id = feed_url.split("urn:li:activity:")[1].split("/")[0]
        return f"https://www.linkedin.com/analytics/post-summary/urn:li:activity:{act_id}/"
    print("URL 형식이 잘못되었습니다. 예) urn:li:activity:1234567890")
    return None

def get_next_row_index() -> int:
    rng = f"{SHEET_NAME}!C4:C"
    rows = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID, range=rng, majorDimension='ROWS'
    ).execute().get('values', [])
    return 4 + len(rows)

# ------------------------------------------------
# 4. Selenium 웹드라이버 (로컬 + CI 공통)
# ------------------------------------------------
def init_driver(download_dir: str) -> webdriver.Chrome:
    chrome_options = Options()
    # CI 환경(Linux)에서만 chromium-browser 사용
    if platform.system() == "Linux":
        chrome_options.binary_location = "/usr/bin/google-chrome"

    # 헤드리스 모드 (새 Chrome 헤드리스)
    chrome_options.add_argument("--headless=new")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--window-size=1920,1080")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--remote-debugging-port=9222")  # 추가
    chrome_options.add_argument("--disable-extensions")  # 추가
    # UI 언어 고정이 필요하면 주석 해제
    # chrome_options.add_argument("--lang=ko-KR")

    chrome_options.add_experimental_option('excludeSwitches', ['enable-automation'])
    chrome_options.add_experimental_option('useAutomationExtension', False)
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_argument(
        "--user-agent=Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    )

    prefs = {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
    }
    chrome_options.add_experimental_option("prefs", prefs)

    service_obj = Service(ChromeDriverManager().install())
    return webdriver.Chrome(service=service_obj, options=chrome_options)

# ------------------------------------------------
# 5. LinkedIn 로그인
# ------------------------------------------------
def login_linkedin(driver: webdriver.Chrome) -> bool:
    email, pwd = get_linkedin_credentials()
    if not email or not pwd:
        print("[ERROR] LinkedIn 로그인 정보가 없습니다.")
        return False

    driver.get("https://www.linkedin.com/login")
    time.sleep(2)
    try:
        driver.find_element(By.ID, "username").send_keys(email)
        driver.find_element(By.ID, "password").send_keys(pwd)
        driver.find_element(By.XPATH, "//button[@type='submit']").click()
    except Exception as e:
        print("로그인 오류:", e)
        return False

    time.sleep(3)
    return True

# ------------------------------------------------
# 6. Analytics XLSX 다운로드 + 파싱
# ------------------------------------------------
def download_xlsx(driver, wait_sec: int = 60) -> bool:
    wait = WebDriverWait(driver, wait_sec)

    # 가능한 모든 셀렉터를 차례로 시도
    selectors = [
        # ① 한글 UI: “다운로드” 텍스트 + primary 버튼
        (By.XPATH,
         "//button[contains(@class,'artdeco-button--primary') and .//span[contains(normalize-space(.),'다운로드')]]"
        ),
        # ② 아이콘 기준 (data-test-icon)
        (By.XPATH,
         "//button[.//svg[@data-test-icon='download-small']]"
        ),
        # ③ 영어 UI: “Download” 텍스트 + primary 버튼
        (By.XPATH,
         "//button[contains(@class,'artdeco-button--primary') and .//span[contains(translate(normalize-space(.),'DOWNLOAD','download'),'download')]]"
        ),
        # ④ aria‑label 속성 (영문)
        (By.XPATH,
         "//button[@aria-label='Download' or @aria-label='다운로드']"
        ),
        # ⑤ CSS: data-control-name (만약 존재한다면)
        (By.CSS_SELECTOR, "button[data-control-name='download_post_analytics']")
    ]

    btn = None
    for by, sel in selectors:
        try:
            btn = wait.until(EC.element_to_be_clickable((by, sel)))
            print(f"[DEBUG] selector matched: {by} {sel}")
            break
        except Exception:
            continue

    if not btn:
        print("[ERROR] 다운로드 버튼을 찾을 수 없습니다 (모든 셀렉터 실패).")
        # 디버깅: 스크린샷 남기기
        driver.save_screenshot("ci_fail_download.png")
        return False

    # 화면 중앙으로 스크롤
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
    time.sleep(1)

    # 클릭 시도
    try:
        btn.click()
    except Exception:
        driver.execute_script("arguments[0].click();", btn)

    # 충분한 시간 대기
    time.sleep(10)
    return True

def get_latest_xlsx(download_dir: str) -> str | None:
    files = glob.glob(os.path.join(download_dir, "*.xlsx"))
    return max(files, key=os.path.getctime) if files else None

def parse_date_time_strings(date_str: str, time_str: str) -> str:
    m = re.match(r"(\d{4})년\s*(\d{1,2})월\s*(\d{1,2})일", date_str.strip())
    y, mth, d = map(int, m.groups()) if m else (1970, 1, 1)

    is_pm = "오후" in time_str
    hour, minute = map(int, time_str.replace("오전", "").replace("오후", "").strip().split(":"))
    if is_pm and hour < 12:
        hour += 12
    if not is_pm and hour == 12:
        hour = 0
    dt = datetime.datetime(y, mth, d, hour, minute)
    return dt.strftime("%Y-%m-%d %H:%M:%S")

def parse_excel(path: str):
    mapping = {
        "impression": "exposure", "노출": "exposure",
        "members reached": "reached", "회원 도달": "reached",
        "reactions": "reactions", "반응": "reactions",
        "comments": "comments", "댓글": "comments",
        "reposts": "reposts", "퍼감": "reposts",
        "게시일": "post_date", "게시 시간": "post_time",
    }
    mapping = {k.lower(): v for k, v in mapping.items()}
    metrics = {v: None for v in mapping.values()}

    df = pd.read_excel(path, sheet_name="실적", header=None)
    for _, row in df.iterrows():
        if pd.isna(row[0]):
            continue
        key = str(row[0]).strip().lower()
        if key in mapping:
            metrics[mapping[key]] = str(row[1]).strip() if not pd.isna(row[1]) else ""

    for k in ["exposure", "reached", "reactions", "comments", "reposts"]:
        try:
            metrics[k] = float(metrics[k]) if metrics[k] else 0
        except Exception:
            metrics[k] = 0

    if metrics["post_date"] and metrics["post_time"]:
        post_time = parse_date_time_strings(metrics["post_date"], metrics["post_time"])
    else:
        post_time = (datetime.datetime.utcnow() + datetime.timedelta(hours=9)) \
                        .strftime("%Y-%m-%d %H:%M:%S")

    return (
        metrics["exposure"], metrics["reached"], metrics["reactions"],
        metrics["comments"], metrics["reposts"], post_time
    )

# ------------------------------------------------
# 7. 시트 기록
# ------------------------------------------------
def write_metrics_to_sheet(exposure, reached, reactions, comments, reposts, row_idx: int):
    rng = f"{SHEET_NAME}!C{row_idx}:G{row_idx}"
    service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID, range=rng, valueInputOption='USER_ENTERED',
        body={'values': [[exposure, reached, reactions, comments, reposts]]}
    ).execute()

def write_post_time_to_sheet(post_time: str):
    service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID, range=f"{SHEET_NAME}!G2",
        valueInputOption='USER_ENTERED', body={'values': [[post_time]]}
    ).execute()

# ------------------------------------------------
# 8. 메인
# ------------------------------------------------
def main():
    dl_dir = os.path.join(os.path.expanduser("~"), "Downloads")
    driver = init_driver(dl_dir)

    if not login_linkedin(driver):
        driver.quit()
        sys.exit(1)
    print("자동 로그인 성공")

    url = get_analytics_url()
    if not url:
        driver.quit()
        sys.exit(1)
    print("[INFO] Analytics URL:", url)

    driver.get(url)
    # CI 환경이 느릴 수 있으므로 충분히 대기
    time.sleep(30)

    if not download_xlsx(driver):
        driver.quit()
        sys.exit(1)

    xlsx = get_latest_xlsx(dl_dir)
    if not xlsx:
        print("XLSX 파일을 찾지 못했습니다.")
        driver.quit()
        sys.exit(1)

    print("[INFO] 파일 경로:", xlsx)
    exposure, reached, reactions, comments, reposts, post_time = parse_excel(xlsx)
    row = get_next_row_index()
    write_metrics_to_sheet(exposure, reached, reactions, comments, reposts, row)
    write_post_time_to_sheet(post_time)
    print(f"[INFO] 시트 기록 완료 (행 {row})")

    try:
        os.remove(xlsx)
    except Exception as e:
        print("임시 파일 삭제 실패:", e)

    driver.quit()
    print("[INFO] 작업 완료")

if __name__ == "__main__":
    main()
