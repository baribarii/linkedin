import os
import time
import datetime
import glob
import re
import json
import pandas as pd

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

from googleapiclient.discovery import build
from google.oauth2.service_account import Credentials

# ====================================
# 1. 구글 스프레드시트 API 인증 (Secrets)
# ====================================
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# GitHub Secrets로부터 Plain JSON 형태로 받은 경우
google_sheet_api_json = os.getenv("LINKEDIN_GOOGLESHEET_API")
if not google_sheet_api_json:
    raise EnvironmentError("환경변수 LINKEDIN_GOOGLESHEET_API가 설정되지 않았습니다.")

# 파일로 저장
SERVICE_ACCOUNT_FILE = "service_account_temp.json"
with open(SERVICE_ACCOUNT_FILE, "w", encoding="utf-8") as f:
    f.write(google_sheet_api_json)

creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)

# 스프레드시트 정보
SPREADSHEET_ID = '1fQTqTrNGwSNGi9EzyK8A2ZqU48IbXG-YrL2ImhXm74w'
SHEET_NAME = '시트4'

service = build('sheets', 'v4', credentials=creds)

def get_analytics_url():
    """
    시트 내 C2 셀에 입력된 feed URL을 읽어,
    LinkedIn Analytics URL로 변환한 뒤 반환한다.
    """
    sheet_range = f"{SHEET_NAME}!C2"
    result = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=sheet_range
    ).execute()
    values = result.get("values", [])
    if not values:
        print("C2 셀의 값이 없습니다.")
        return None

    feed_url = values[0][0].strip()
    try:
        if "urn:li:activity:" in feed_url:
            id_part = feed_url.split("urn:li:activity:")[1].split("/")[0]
        else:
            print("C2 셀에서 share/activity ID 추출 실패. URL 형식 확인 필요.")
            return None
        analytics_url = f"https://www.linkedin.com/analytics/post-summary/urn:li:activity:{id_part}/"
        return analytics_url
    except Exception as e:
        print("Analytics URL 생성 중 에러:", e)
        return None

def get_next_row_index():
    read_range = f"{SHEET_NAME}!C4:C"
    result = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=read_range,
        majorDimension='ROWS'
    ).execute()

    rows = result.get('values', [])
    used_count = len(rows)
    return 4 + used_count

# ====================================
# 3. Selenium 웹드라이버 설정
# ====================================
def init_driver(download_dir):
    chrome_options = Options()
    chrome_options.add_argument(
        "--user-agent=Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/99.0.4844.74 Safari/537.36"
    )
    chrome_options.add_experimental_option('excludeSwitches', ['enable-automation'])
    chrome_options.add_experimental_option('useAutomationExtension', False)
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    # GitHub Actions 같은 CI 환경에서는 헤드리스 모드가 필요할 수 있음
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")

    prefs = {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    }
    chrome_options.add_experimental_option("prefs", prefs)

    service_obj = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service_obj, options=chrome_options)
    return driver

# ====================================
# 4. LinkedIn 자동 로그인 (Secrets → env)
# ====================================
def login_linkedin(driver):
    linkedin_email = os.getenv("LINKEDIN_EMAIL")
    linkedin_password = os.getenv("LINKEDIN_PASSWORD")

    if not linkedin_email or not linkedin_password:
        print("[ERROR] LINKEDIN_EMAIL 또는 LINKEDIN_PASSWORD 환경변수가 설정되지 않았습니다.")
        return False

    driver.get("https://www.linkedin.com/login")
    time.sleep(2)

    try:
        email_input = driver.find_element(By.ID, "username")
        password_input = driver.find_element(By.ID, "password")
    except Exception as e:
        print("로그인 입력창을 찾지 못했습니다:", e)
        return False

    email_input.send_keys(linkedin_email)
    password_input.send_keys(linkedin_password)

    try:
        login_button = driver.find_element(By.XPATH, "//button[@type='submit']")
        login_button.click()
    except Exception as e:
        print("로그인 버튼 클릭 중 에러:", e)
        return False

    time.sleep(3)
    return True

# ====================================
# 5. Analytics 페이지 XLSX 다운로드
# ====================================
def download_xlsx(driver, download_wait=10):
    try:
        download_button = driver.find_element(By.XPATH, "//button[contains(., 'Download') or contains(., '다운로드')]")
        download_button.click()
    except Exception as e:
        print("[ERROR] 다운로드 버튼을 찾지 못했습니다:", e)
        return False
    time.sleep(download_wait)
    return True

def get_latest_xlsx(download_dir):
    files = glob.glob(os.path.join(download_dir, "*.xlsx"))
    if not files:
        return None
    latest_file = max(files, key=os.path.getctime)
    return latest_file

def parse_date_time_strings(date_str, time_str):
    m = re.match(r"(\d{4})년\s*(\d{1,2})월\s*(\d{1,2})일", date_str.strip())
    if m:
        year, month, day = map(int, m.groups())
    else:
        year, month, day = 1970, 1, 1

    is_am = "오전" in time_str
    is_pm = "오후" in time_str
    t = time_str.replace("오전", "").replace("오후", "").strip()
    hour, minute = 0, 0
    if ":" in t:
        hour_str, minute_str = t.split(":")
        hour = int(hour_str)
        minute = int(minute_str)
    if is_pm and hour < 12:
        hour += 12
    if is_am and hour == 12:
        hour = 0

    dt = datetime.datetime(year, month, day, hour, minute, 0)
    return dt.strftime("%Y-%m-%d %H:%M:%S")

def parse_excel(file_path):
    mapping = {
        "impression": "exposure",
        "노출": "exposure",
        "members reached": "reached",
        "회원 도달": "reached",
        "reactions": "reactions",
        "반응": "reactions",
        "comments": "comments",
        "댓글": "comments",
        "reposts": "reposts",
        "퍼감": "reposts",
        "게시일": "post_date",
        "게시 시간": "post_time"
    }
    mapping = {k.lower(): v for k, v in mapping.items()}

    metrics = {
        "exposure": None,
        "reached": None,
        "reactions": None,
        "comments": None,
        "reposts": None,
        "post_date": None,
        "post_time": None
    }
    try:
        df = pd.read_excel(file_path, sheet_name="실적", header=None)
    except Exception as e:
        print("엑셀 파일 읽기 에러:", e)
        return None

    for idx, row in df.iterrows():
        if pd.isna(row[0]):
            continue
        key_val = str(row[0]).strip().lower()
        if key_val in mapping:
            metric_key = mapping[key_val]
            try:
                value = str(row[1]).strip() if not pd.isna(row[1]) else ""
            except Exception:
                value = ""
            metrics[metric_key] = value

    # 수치 메트릭 변환
    for k in ["exposure", "reached", "reactions", "comments", "reposts"]:
        try:
            if not metrics[k]:
                metrics[k] = 0
            else:
                metrics[k] = float(metrics[k])
        except Exception:
            metrics[k] = 0

    if metrics["post_date"] and metrics["post_time"]:
        post_time_str = parse_date_time_strings(metrics["post_date"], metrics["post_time"])
    else:
        pt = datetime.datetime.utcnow() + datetime.timedelta(hours=9)
        post_time_str = pt.strftime("%Y-%m-%d %H:%M:%S")

    return (
        metrics["exposure"],
        metrics["reached"],
        metrics["reactions"],
        metrics["comments"],
        metrics["reposts"],
        post_time_str
    )

def write_metrics_to_sheet(exposure, reached, reactions, comments, reposts, row_index):
    values = [[exposure, reached, reactions, comments, reposts]]
    range_data = f"{SHEET_NAME}!C{row_index}:G{row_index}"
    service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=range_data,
        valueInputOption='USER_ENTERED',
        body={'values': values}
    ).execute()

def write_post_time_to_sheet(post_time_str):
    range_data = f"{SHEET_NAME}!G2"
    service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=range_data,
        valueInputOption='USER_ENTERED',
        body={'values': [[post_time_str]]}
    ).execute()

def main():
    download_dir = os.path.join(os.path.expanduser("~"), "Downloads")
    driver = init_driver(download_dir)

    if not login_linkedin(driver):
        print("자동 로그인 실패.")
        driver.quit()
        return
    print("자동 로그인 성공.")

    analytics_url = get_analytics_url()
    if analytics_url is None:
        print("Analytics URL 생성 실패.")
        driver.quit()
        return
    print("[INFO] Analytics URL:", analytics_url)

    driver.get(analytics_url)
    time.sleep(5)

    if not download_xlsx(driver, download_wait=10):
        print("다운로드 실패.")
        driver.quit()
        return

    file_path = get_latest_xlsx(download_dir)
    if not file_path:
        print("다운로드된 XLSX 파일을 찾을 수 없습니다.")
        driver.quit()
        return
    print("[INFO] 다운로드 파일 경로:", file_path)

    parsed = parse_excel(file_path)
    if not parsed:
        print("엑셀 파싱 중 에러.")
        driver.quit()
        return

    exposure, reached, reactions, comments, reposts, post_time_str = parsed
    print(f"[INFO] 추출 데이터: 노출={exposure}, 회원 도달={reached}, "
          f"반응={reactions}, 댓글={comments}, 퍼감={reposts}, 게시시간={post_time_str}")

    row_index = get_next_row_index()
    print(f"[INFO] 데이터 작성 행: {row_index}")

    write_metrics_to_sheet(exposure, reached, reactions, comments, reposts, row_index)
    print(f"[INFO] 시트 {row_index}행 메트릭 기록 완료.")

    write_post_time_to_sheet(post_time_str)
    print(f"[INFO] G2 셀에 게시시간({post_time_str}) 기록 완료.")

    try:
        os.remove(file_path)
    except Exception as e:
        print("파일 삭제 에러:", e)

    driver.quit()
    print("[INFO] 작업 완료.")

if __name__ == "__main__":
    main()
