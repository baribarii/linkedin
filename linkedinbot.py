import os
import time
import datetime
import glob
import re
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
# 1. Google Sheets API 설정
# ====================================
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SERVICE_ACCOUNT_FILE = 'compelling-pen-456906-q0-f280b92105f7.json'
creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)

# 스프레드시트 ID와 시트 이름
SPREADSHEET_ID = '1fQTqTrNGwSNGi9EzyK8A2ZqU48IbXG-YrL2ImhXm74w'
SHEET_NAME = '시트4'
service = build('sheets', 'v4', credentials=creds)

# ====================================
# 2. 스프레드시트에서 C2 셀의 feed URL을 읽어 Analytics URL 생성
# ====================================
def get_analytics_url():
    # C2 셀의 값 읽기
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
    # 예: https://www.linkedin.com/feed/update/urn:li:activity:{id}/
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
    # 자동화 특성을 낮추기 위한 옵션
    chrome_options.add_argument(
        "--user-agent=Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/99.0.4844.74 Safari/537.36"
    )
    chrome_options.add_experimental_option('excludeSwitches', ['enable-automation'])
    chrome_options.add_experimental_option('useAutomationExtension', False)
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    
    # 다운로드 폴더 설정
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
    # Keychain에서 "LinkedIn" 서비스로 저장된 정보를 불러옴
    linkedin_email = keyring.get_password("LinkedIn", "email")
    linkedin_password = keyring.get_password("LinkedIn", "password")
    
    if not linkedin_email or not linkedin_password:
        print("Keychain에서 LinkedIn 로그인 정보를 찾을 수 없습니다.")
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
# 5. XLSX 파일 다운로드 함수 (Analytics 페이지에서)
# ====================================
def download_xlsx(driver, download_wait=10):
    try:
        download_button = driver.find_element(By.XPATH, "//button[contains(., 'Download') or contains(., '다운로드')]")
        download_button.click()
    except Exception as e:
        print("다운로드 버튼을 찾지 못했습니다:", e)
        return False
    time.sleep(download_wait)
    return True

# ====================================
# 6. 최근에 다운로드된 XLSX 파일 찾기
# ====================================
def get_latest_xlsx(download_dir):
    list_of_files = glob.glob(os.path.join(download_dir, "*.xlsx"))
    if not list_of_files:
        return None
    latest_file = max(list_of_files, key=os.path.getctime)
    return latest_file

# ====================================
# 7. 날짜 및 시간 문자열을 합쳐서 변환하는 함수
# ====================================
def parse_date_time_strings(date_str, time_str):
    """
    date_str: 예, "2025년 4월 1일"
    time_str: 예, "오전 6:26" 또는 "오후 3:15"
    return: "YYYY-MM-DD HH:MM:SS"
    """
    # 1) 날짜 파싱
    m = re.match(r"(\d{4})년\s*(\d{1,2})월\s*(\d{1,2})일", date_str.strip())
    if m:
        year, month, day = map(int, m.groups())
    else:
        year, month, day = 1970, 1, 1

    # 2) 시간 파싱
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
# 8. XLSX 파일 파싱 함수 (실적 시트)
# ====================================
def parse_excel(file_path):
    """
    '실적' 시트에서 각 행의 A열(메트릭 이름)을 읽어 해당 B열의 값을 추출한다.
    
    지원 메트릭:
      - "impression" 또는 "노출"          → exposure
      - "members reached" 또는 "회원 도달"  → reached
      - "reactions" 또는 "반응"            → reactions
      - "comments" 또는 "댓글"             → comments
      - "reposts" 또는 "퍼감"              → reposts
      - "게시일"                           → post_date
      - "게시시간"                         → post_time
      
    게시일과 게시시간 값을 모두 읽어 있을 경우 결합하여 최종 게시시간 문자열("YYYY-MM-DD HH:MM:SS")을 생성한다.
    만약 게시일 또는 게시시간이 없으면, 현재 시간을 사용한다.
    """
    # 매핑: 두 개의 게시 관련 키를 분리하여 설정
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
    # 초기값: 게시일과 게시시간은 각각 None, 기타 메트릭은 None
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
        print("엑셀 파일 읽기 중 에러:", e)
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

    # 나머지 메트릭은 수치로 변환 (게시일/게시시간 제외)
    for k in ["exposure", "reached", "reactions", "comments", "reposts"]:
        try:
            # 값이 문자열이면 float 변환, 아니면 0
            if metrics[k] == "" or metrics[k] is None:
                metrics[k] = 0
            else:
                metrics[k] = float(metrics[k])
        except Exception:
            metrics[k] = 0

    # 게시일과 게시시간을 결합
    if metrics["post_date"] and metrics["post_time"]:
        post_time_str = parse_date_time_strings(metrics["post_date"], metrics["post_time"])

    else:
        # 둘 중 하나라도 없으면 현재 KST 기준 시간 사용
        pt = datetime.datetime.utcnow() + datetime.timedelta(hours=9)
        post_time_str = pt.strftime("%Y-%m-%d %H:%M:%S")
    
    return (metrics["exposure"], metrics["reached"], metrics["reactions"],
            metrics["comments"], metrics["reposts"], post_time_str)

# ====================================
# 9. 구글 스프레드시트에 기록 함수
# ====================================
def write_metrics_to_sheet(exposure, reached, reactions, comments, reposts, row_index):
    # 메트릭 5개를 C~G열에 기록
    values = [[exposure, reached, reactions, comments, reposts]]
    range_data = f"{SHEET_NAME}!C{row_index}:G{row_index}"
    service.spreadsheets().values().update(
         spreadsheetId=SPREADSHEET_ID,
         range=range_data,
         valueInputOption='USER_ENTERED',
         body={'values': values}
    ).execute()

def write_post_time_to_sheet(post_time_str):
    # 게시시간을 G2 셀에 덮어쓰기 (형식 "YYYY-MM-DD HH:mm:ss")
    range_data = f"{SHEET_NAME}!G2"
    service.spreadsheets().values().update(
         spreadsheetId=SPREADSHEET_ID,
         range=range_data,
         valueInputOption='USER_ENTERED',
         body={'values': [[post_time_str]]}
    ).execute()

# ====================================
# 10. 메인 실행 함수
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
        print("Analytics URL을 생성하지 못했습니다.")
        driver.quit()
        return
    print("생성된 Analytics URL:", analytics_url)
    
    driver.get(analytics_url)
    time.sleep(5)  # 페이지 로드 대기
    
    if not download_xlsx(driver, download_wait=10):
        print("다운로드 실패.")
        driver.quit()
        return
    
    file_path = get_latest_xlsx(download_dir)
    if not file_path:
        print("다운로드한 XLSX 파일을 찾을 수 없습니다.")
        driver.quit()
        return
    print("다운로드 파일 경로:", file_path)
    
    parsed = parse_excel(file_path)
    if not parsed:
        print("엑셀 파싱 중 에러 발생.")
        driver.quit()
        return
    exposure, reached, reactions, comments, reposts, post_time_str = parsed
    print(f"추출 데이터: 노출={exposure}, 회원 도달={reached}, 반응={reactions}, 댓글={comments}, 퍼감={reposts}, 게시시간={post_time_str}")
    
    write_metrics_to_sheet(exposure, reached, reactions, comments, reposts, row_index=4)
    print("[스프레드시트] 메트릭 기록 완료.")
    
    write_post_time_to_sheet(post_time_str)
    print(f"[스프레드시트] G2 셀에 게시시간({post_time_str}) 기록 완료.")
    
    try:
        os.remove(file_path)
    except Exception as e:
        print("XLSX 파일 삭제 중 에러:", e)
    
    driver.quit()
    print("작업 완료.")
    
if __name__ == "__main__":
    main()

