import os
import time
import datetime
import glob
import re
import json
import base64
import pandas as pd
import keyring  # macOS Keychain 활용

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

from googleapiclient.discovery import build
from google.oauth2.service_account import Credentials

# ====================================
# 1. 구글 스프레드시트 API 인증 설정
# ====================================
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# 환경변수 이름을 'LINKEDIN_GOOGLESHEET_API'로 사용 (Base64 인코딩된 서비스 계정 JSON 문자열)
encoded_creds = os.getenv("LINKEDIN_GOOGLESHEET_API")
if encoded_creds:
    try:
        # Base64 디코딩하여 문자열 얻기
        json_creds = base64.b64decode(encoded_creds).decode("utf-8")
    except Exception as e:
        raise ValueError("Base64 디코딩에 실패했습니다: " + str(e))
    
    # 임시 파일로 저장
    with open("service_account_temp.json", "w") as f:
        f.write(json_creds)
    SERVICE_ACCOUNT_FILE = "service_account_temp.json"
    print("[INFO] SERVICE_ACCOUNT_FILE created from env LINKEDIN_GOOGLESHEET_API.")
else:
    # 로컬 실행 시 환경변수가 없다면, 로컬 파일을 사용
    SERVICE_ACCOUNT_FILE = "compelling-pen-456906-q0-f280b92105f7.json"
    print("[INFO] Using local service account file:", SERVICE_ACCOUNT_FILE)

if not SERVICE_ACCOUNT_FILE or not os.path.exists(SERVICE_ACCOUNT_FILE):
    raise FileNotFoundError("Service account JSON is not provided or file not found.")

creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)

# 스프레드시트 ID와 시트 이름 (요청하신 값)
SPREADSHEET_ID = '1fQTqTrNGwSNGi9EzyK8A2ZqU48IbXG-YrL2ImhXm74w'
SHEET_NAME = '시트4'

service = build('sheets', 'v4', credentials=creds)

# ====================================
# 2. 스프레드시트에서 C2 셀의 feed URL을 읽어 Analytics URL 생성
# ====================================
def get_analytics_url():
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
            print("C2 셀의 URL 형식이 올바르지 않습니다.")
            return None
        analytics_url = f"https://www.linkedin.com/analytics/post-summary/urn:li:activity:{id_part}/"
        return analytics_url
    except Exception as e:
        print("id 추출 중 에러 발생:", e)
        return None

# ====================================
# 3. Selenium 웹드라이버 설정 (Chrome)
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
    # headless 모드를 원하면 주석 해제: chrome_options.add_argument("--headless")
    
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
# 4. macOS Keychain을 통한 LinkedIn 자동 로그인
# ====================================
def login_linkedin(driver):
    linkedin_email = keyring.get_password("LinkedIn", "email")
    linkedin_password = keyring.get_password("LinkedIn", "password")
    
    if not linkedin_email or not linkedin_password:
        print("[ERROR] Keychain에서 LinkedIn 로그인 정보를 찾을 수 없습니다.")
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
# 5. Analytics 페이지에서 XLSX 파일 다운로드
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

# ====================================
# 6. 최근에 다운로드된 XLSX 파일 찾기
# ====================================
def get_latest_xlsx(download_dir):
    files = glob.glob(os.path.join(download_dir, "*.xlsx"))
    if not files:
        return None
    return max(files, key=os.path.getctime)

# ====================================
# 7. 게시일과 게시시간 결합 후 날짜 문자열 생성
# ====================================
def parse_date_time_strings(date_str, time_str):
    m = re.match(r"(\d{4})년\s*(\d{1,2})월\s*(\d{1,2})일", date_str.strip())
    if m:
        year, month, day = map(int, m.groups())
    else:
        year, month, day = 1970, 1, 1

    is_am = "오전" in time_str
    is_pm = "오후" in time_str
    t = time_str.replace("오전", "").replace("오후", "").strip()
    if ":" in t:
        hour_str, minute_str = t.split(":")
        hour = int(hour_str)
        minute = int(minute_str)
    else:
        hour, minute = 0, 0
    if is_pm and hour < 12:
        hour += 12
    if is_am and hour == 12:
        hour = 0

    dt = datetime.datetime(year, month, day, hour, minute, 0)
    return dt.strftime("%Y-%m-%d %H:%M:%S")

# ====================================
# 8. XLSX 파일 파싱 (실적 시트)
# ====================================
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
    mapping = { k.lower(): v for k, v in mapping.items() }
    
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

    for k in ["exposure", "reached", "reactions", "comments", "reposts"]:
        try:
            if metrics[k] == "" or metrics[k] is None:
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
    
    return (metrics["exposure"], metrics["reached"], metrics["reactions"],
            metrics["comments"], metrics["reposts"], post_time_str)

# ====================================
# 9. 구글 스프레드시트 업데이트
# ====================================
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

# ====================================
# 10. 메인 실행 함수 (한 회차 실행)
# ====================================
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
    print(f"[INFO] 추출 데이터: 노출={exposure}, 회원 도달={reached}, 반응={reactions}, "
          f"댓글={comments}, 퍼감={reposts}, 게시시간={post_time_str}")
    
    row_index = 4
    write_metrics_to_sheet(exposure, reached, reactions, comments, reposts, row_index)
    print(f"[INFO] {SHEET_NAME} 시트 {row_index}행에 메트릭 기록 완료.")
    
    write_post_time_to_sheet(post_time_str)
    print(f"[INFO] G2 셀에 게시시간({post_time_str}) 기록 완료.")
    
    try:
        os.remove(file_path)
    except Exception as e:
        print("XLSX 파일 삭제 에러:", e)
    
    driver.quit()
    print("\n[INFO] 작업 완료.")

if __name__ == "__main__":
    main()
