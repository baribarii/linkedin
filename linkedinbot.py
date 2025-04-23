#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
LinkedIn 포스트 데이터를 주기적으로 Google 스프레드시트에 기록
로컬(macOS): 키체인 'LinkedIn' 항목(email / password) 사용
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
import pickle

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
# 1. Google Service Account 인증
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
# 4. Selenium 웹드라이버 (로컬 + CI 공통) - 개선됨
# ------------------------------------------------
def init_driver(download_dir: str) -> webdriver.Chrome:
    import random  # 함수 상단에 추가
    
    chrome_options = Options()
    # CI 환경(Linux)에서만 chromium-browser 사용
    if platform.system() == "Linux":
        chrome_options.binary_location = "/usr/bin/google-chrome"

    # 헤드리스 모드 개선
    chrome_options.add_argument("--headless=new")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--start-maximized")  # 창 최대화
    chrome_options.add_argument("--window-size=1920,1080")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--remote-debugging-port=9222")
    chrome_options.add_argument("--disable-extensions")
    
    # 봇 탐지 방지 추가 설정
    chrome_options.add_argument("--disable-blink-features")
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    
    # 쿠키 및 캐시 활성화
    chrome_options.add_argument("--enable-cookies")
    chrome_options.add_argument("--profile-directory=Default")
    
    # 프록시 설정 추가
    proxy = os.getenv("HTTPS_PROXY") or os.getenv("HTTP_PROXY")
    if proxy:
        print(f"[INFO] 프록시 사용: {proxy}")
        chrome_options.add_argument(f'--proxy-server={proxy}')
    
    # User Agent 랜덤화
    user_agents = [
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36",
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36",
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
        "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36"
    ]
    chosen_user_agent = random.choice(user_agents)
    print(f"[INFO] 선택된 User-Agent: {chosen_user_agent}")
    chrome_options.add_argument(f'--user-agent={chosen_user_agent}')
    
    # 추가 실행 옵션
    chrome_options.add_argument("--lang=en-US,en;q=0.9")
    chrome_options.add_experimental_option('excludeSwitches', ['enable-automation'])
    chrome_options.add_experimental_option('useAutomationExtension', False)
    
    prefs = {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
    }
    chrome_options.add_experimental_option("prefs", prefs)

    service_obj = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service_obj, options=chrome_options)
    
    # WebDriver 속성 마스킹
    driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
    
    return driver

# ------------------------------------------------
# 5. LinkedIn 로그인 (기존 함수 유지)
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
# 6. 로그인 후 인증 확인 (새 함수 추가)
# ------------------------------------------------
def handle_login_verification(driver):
    """로그인 후 추가 인증 또는 보안 확인 페이지 처리"""
    time.sleep(5)  # 페이지 로드 대기
    
    # CAPTCHA 또는 보안 확인 감지
    security_prompts = [
        "security verification", "보안 확인", "verify it's you", 
        "자동화된 접근", "automated access", "captcha", "퍼즐",
        "unusual login", "비정상적인 로그인"
    ]
    
    page_source = driver.page_source.lower()
    
    # 보안 확인 페이지 감지
    for prompt in security_prompts:
        if prompt in page_source:
            print(f"[WARN] 보안 확인 감지: '{prompt}'")
            driver.save_screenshot("security_challenge.png")
            return False
    
    # 현재 URL 확인
    current_url = driver.current_url
    if "checkpoint" in current_url or "security-verification" in current_url:
        print(f"[WARN] 보안 검증 URL 감지: {current_url}")
        driver.save_screenshot("security_url.png")
        return False
    
    # LinkedIn 홈페이지 확인
    if "/feed" in current_url:
        print("[INFO] LinkedIn 홈 페이지 확인됨")
        return True
    
    return True

# ------------------------------------------------
# 7. Analytics 페이지 대기 (새 함수 추가)
# ------------------------------------------------
def wait_for_analytics_page(driver, timeout=90):
    """Analytics 페이지가 로드될 때까지 기다립니다."""
    print("[INFO] Analytics 페이지 로드 대기...")
    
    try:
        # 더 긴 타임아웃과 명시적 대기 조건
        WebDriverWait(driver, timeout).until(
            lambda d: d.execute_script("return document.readyState") == "complete"
        )
        
        # 더 넓은 범위의 선택자 사용
        selectors = [
            "main.scaffold-layout__main", 
            "div.scaffold-layout__main",
            "div[data-test-id='post-analytics']",
            "div[data-control-name='analytics']",
            "div.analytics",
            "section.insights-module"
        ]
        
        for selector in selectors:
            try:
                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, selector))
                )
                print(f"[INFO] 발견된 Analytics 요소: {selector}")
                return True
            except:
                pass
        
        # JavaScript로 페이지 상태 확인
        js_result = driver.execute_script("""
            return {
                title: document.title,
                url: window.location.href,
                hasLoginForm: !!document.querySelector('form#login'),
                hasAnalyticsContent: !!document.querySelector('.analytics, [data-control-name*="analytics"]'),
                bodyText: document.body.innerText.substring(0, 200)
            }
        """)
        
        print(f"[DEBUG] 페이지 상태: {js_result}")
        
        # 페이지 소스 저장 (디버깅용)
        with open("page_source.html", "w", encoding="utf-8") as f:
            f.write(driver.page_source)
        
        return js_result.get('hasAnalyticsContent', False)
        
    except Exception as e:
        print(f"[ERROR] Analytics 페이지 대기 실패: {e}")
        return False

# ------------------------------------------------
# 8. 페이지 로드 대기 (기존 함수 유지)
# ------------------------------------------------
def wait_for_page_load(driver, timeout=60):
    """페이지가 완전히 로드될 때까지 기다립니다."""
    print("[INFO] 페이지 로딩 대기 시작...")
    
    # 먼저 document.readyState 확인
    try:
        WebDriverWait(driver, timeout).until(
            lambda d: d.execute_script("return document.readyState") == "complete"
        )
        print("[INFO] 문서 로드 완료 (readyState: complete)")
    except Exception as e:
        print(f"[WARN] 문서 로드 대기 실패: {e}")
    
    # 추가 대기 (AJAX 완료를 위해)
    time.sleep(5)
    
    # LinkedIn 페이지 특정 요소 확인
    try:
        WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "main.scaffold-layout__main, div.scaffold-layout__main"))
        )
        print("[INFO] LinkedIn 메인 컨테이너 감지됨")
    except Exception as e:
        print(f"[WARN] LinkedIn 메인 컨테이너 감지 실패: {e}")
    
    # 추가 스크롤 시도 (AJAX 콘텐츠 로드 유도)
    try:
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight/2);")
        time.sleep(2)
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(2)
        driver.execute_script("window.scrollTo(0, 0);")
        print("[INFO] 페이지 스크롤 수행")
    except Exception as e:
        print(f"[WARN] 스크롤 시도 실패: {e}")
    
    # 최종 대기
    time.sleep(5)
    
    return True

# ------------------------------------------------
# 9. 페이지 구조 분석 (기존 함수 유지)
# ------------------------------------------------
def analyze_page_structure(driver):
    """페이지 구조를 분석하여 디버깅에 도움이 되는 정보를 수집합니다."""
    print("[DEBUG] 페이지 구조 분석 시작...")
    
    # 모든 버튼 요소 찾기
    buttons = driver.find_elements(By.TAG_NAME, "button")
    print(f"[DEBUG] 총 버튼 수: {len(buttons)}")
    
    # 다운로드와 관련된 버튼 찾기
    download_buttons = []
    for btn in buttons[:30]:  # 처음 30개만 확인
        try:
            text = btn.text.strip().lower()
            if "download" in text or "다운로드" in text:
                download_buttons.append({
                    "text": text,
                    "class": btn.get_attribute("class"),
                    "id": btn.get_attribute("id"),
                    "aria-label": btn.get_attribute("aria-label")
                })
        except:
            pass
    
    print(f"[DEBUG] 다운로드 관련 버튼: {len(download_buttons)}")
    for i, btn in enumerate(download_buttons):
        print(f"[DEBUG] 다운로드 버튼 {i+1}: {btn}")
    
    # SVG 아이콘 찾기 (다운로드 아이콘일 수 있음)
    svg_icons = driver.find_elements(By.TAG_NAME, "svg")
    print(f"[DEBUG] SVG 아이콘 수: {len(svg_icons)}")
    
    # 페이지의 주요 컨테이너 확인
    main_containers = driver.find_elements(By.CSS_SELECTOR, "main, section, article")
    print(f"[DEBUG] 주요 컨테이너 수: {len(main_containers)}")
    
    # 특정 패턴의 요소 찾기 (LinkedIn Analytics 관련)
    analytics_elems = driver.find_elements(By.XPATH, "//*[contains(@class, 'analytics') or contains(@id, 'analytics')]")
    print(f"[DEBUG] Analytics 관련 요소 수: {len(analytics_elems)}")
    
    return True

# ------------------------------------------------
# 10. 다운로드 버튼 찾기 (기존 함수 유지)
# ------------------------------------------------
def find_download_button(driver):
    """다양한 방법으로 다운로드 버튼을 찾습니다."""
    print("[INFO] 다운로드 버튼 찾기 시작...")
    
    # 1. 다운로드 관련 텍스트가 있는 모든 요소 찾기
    try:
        xpath_patterns = [
            "//button[contains(translate(., 'DOWNLOAD다운로드', 'download다운로드'), 'download')]",
            "//a[contains(translate(., 'DOWNLOAD다운로드', 'download다운로드'), 'download')]",
            "//*[contains(translate(., 'DOWNLOAD다운로드', 'download다운로드'), 'download') and (self::button or self::a)]"
        ]
        
        for pattern in xpath_patterns:
            elements = driver.find_elements(By.XPATH, pattern)
            if elements:
                print(f"[DEBUG] 패턴 '{pattern}'으로 {len(elements)}개 요소 발견")
                for element in elements[:5]:
                    print(f"[DEBUG] 요소 텍스트: '{element.text}', 클래스: {element.get_attribute('class')}")
                    if element.is_displayed() and element.is_enabled():
                        print("[INFO] 사용 가능한 다운로드 버튼 발견")
                        return element
    except Exception as e:
        print(f"[WARN] 텍스트 기반 버튼 검색 실패: {e}")
    
    # 2. 다운로드 관련 속성을 가진 요소 찾기
    try:
        attribute_patterns = [
            "//*[@aria-label='다운로드' or @aria-label='Download']",
            "//*[contains(@data-control-name, 'download')]",
            "//*[contains(@class, 'download')]"
        ]
        
        for pattern in attribute_patterns:
            elements = driver.find_elements(By.XPATH, pattern)
            if elements:
                print(f"[DEBUG] 속성 패턴 '{pattern}'으로 {len(elements)}개 요소 발견")
                for element in elements[:5]:
                    if element.is_displayed() and element.is_enabled():
                        print("[INFO] 속성 기반 다운로드 버튼 발견")
                        return element
    except Exception as e:
        print(f"[WARN] 속성 기반 버튼 검색 실패: {e}")
    
    # 3. SVG 아이콘을 포함한 요소 찾기
    try:
        svg_parent_patterns = [
            "//button[.//svg[contains(@class, 'download') or @data-test-icon='download-small']]",
            "//a[.//svg[contains(@class, 'download') or @data-test-icon='download-small']]"
        ]
        
        for pattern in svg_parent_patterns:
            elements = driver.find_elements(By.XPATH, pattern)
            if elements:
                print(f"[DEBUG] SVG 패턴 '{pattern}'으로 {len(elements)}개 요소 발견")
                for element in elements[:5]:
                    if element.is_displayed() and element.is_enabled():
                        print("[INFO] SVG 아이콘이 있는 다운로드 버튼 발견")
                        return element
    except Exception as e:
        print(f"[WARN] SVG 기반 버튼 검색 실패: {e}")
    
    print("[WARN] 모든 방법으로 다운로드 버튼을 찾지 못함")
    return None

# ------------------------------------------------
# 11. 다운로드 실행 (기존 함수 유지)
# ------------------------------------------------
def execute_download(driver, download_button=None):
    """다운로드 버튼을 찾고 클릭합니다."""
    if not download_button:
        download_button = find_download_button(driver)

    if download_button:
        try:
            # 요소가 보이도록 스크롤
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", download_button)
            time.sleep(2)
            
            # 클릭 시도
            try:
                download_button.click()
                print("[INFO] 다운로드 버튼 클릭 성공 (직접 클릭)")
            except Exception as e:
                print(f"[WARN] 직접 클릭 실패: {e}")
                driver.execute_script("arguments[0].click();", download_button)
                print("[INFO] 다운로드 버튼 클릭 성공 (JavaScript 클릭)")
            
            # 다운로드 대기
            time.sleep(10)
            return True
        except Exception as e:
            print(f"[ERROR] 다운로드 버튼 클릭 실패: {e}")
    
    # 버튼을 찾지 못한 경우, JavaScript로 시도
    print("[INFO] JavaScript로 다운로드 시도 중...")
    try:
        result = driver.execute_script("""
        function findDownloadButton() {
            // 텍스트로 찾기
            const allElements = document.querySelectorAll('button, a');
            for (const elem of allElements) {
                const text = elem.textContent.toLowerCase();
                if (text.includes('download') || text.includes('다운로드')) {
                    return elem;
                }
            }
            
            // 클래스/속성으로 찾기
            const byClass = document.querySelector('[class*="download"], [aria-label="Download"], [aria-label="다운로드"]');
            if (byClass) return byClass;
            
            // SVG 아이콘으로 찾기
            const withSvg = document.querySelector('button svg[data-test-icon="download-small"]');
            if (withSvg) return withSvg.closest('button');
            
            return null;
        }
        
        const button = findDownloadButton();
        if (button) {
            button.click();
            return true;
        }
        return false;
        """)
        print(f"[DEBUG] JavaScript 다운로드 시도 결과: {result}")
        if result:
            time.sleep(10)
            return True
    except Exception as e:
        print(f"[ERROR] JavaScript 다운로드 시도 실패: {e}")
    
    return False

# ------------------------------------------------
# 12. XLSX 파일 관련 유틸 (기존 함수 유지)
# ------------------------------------------------
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
# 13. 스프레드시트 기록 (기존 함수 유지)
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
# 14. 쿠키 관리 함수 (새 함수 추가)
# ------------------------------------------------
def save_cookies(driver, path="linkedin_cookies.pkl"):
    """현재 세션의 쿠키를 저장합니다."""
    try:
        with open(path, 'wb') as file:
            pickle.dump(driver.get_cookies(), file)
        print(f"[INFO] 쿠키 저장 완료: {path}")
        return True
    except Exception as e:
        print(f"[WARN] 쿠키 저장 실패: {e}")
        return False

def load_cookies(driver, path="linkedin_cookies.pkl"):
    """저장된 쿠키를 로드합니다."""
    try:
        if not os.path.exists(path):
            print(f"[INFO] 쿠키 파일이 없습니다: {path}")
            return False
        
        with open(path, 'rb') as file:
            cookies = pickle.load(file)
        
        # LinkedIn 도메인에 접속 후 쿠키 추가
        driver.get("https://www.linkedin.com")
        for cookie in cookies:
            # 일부 브라우저에서는 expiry 항목이 문제 발생
            if 'expiry' in cookie:
                del cookie['expiry']
            driver.add_cookie(cookie)
        
        print(f"[INFO] 쿠키 로드 완료: {path}")
        return True
    except Exception as e:
        print(f"[WARN] 쿠키 로드 실패: {e}")
        return False

# ------------------------------------------------
# 15. 메인 (수정됨)
# ------------------------------------------------
def main():
    dl_dir = os.path.join(os.path.expanduser("~"), "Downloads")
    
    # 프록시 설정 확인 (GitHub Actions에서 환경변수로 전달됨)
    proxy = os.getenv("HTTPS_PROXY") or os.getenv("HTTP_PROXY")
    if proxy:
        print(f"[INFO] 프록시 설정 감지: {proxy}")
    else:
        print("[WARN] 프록시가 설정되지 않았습니다. LinkedIn 접속이 차단될 수 있습니다.")
    
    try:
        driver = init_driver(dl_dir)
        
        # 쿠키 로드 시도
        cookie_loaded = load_cookies(driver)
        if cookie_loaded:
            driver.get("https://www.linkedin.com/feed/")
            time.sleep(3)
            if "login" not in driver.current_url:
                print("[INFO] 저장된 쿠키로 로그인 성공")
            else:
                print("[INFO] 쿠키 만료, 일반 로그인 시도")
                cookie_loaded = False
        
        # 쿠키 로드 실패 또는 만료 시 일반 로그인
        if not cookie_loaded:
            if not login_linkedin(driver):
                print("[ERROR] LinkedIn 로그인 실패")
                driver.save_screenshot("login_failed.png")
                driver.quit()
                sys.exit(1)
            print("자동 로그인 성공")
            
            # 로그인 성공 시 쿠키 저장
            save_cookies(driver)
        
        # 보안 인증 확인
        if not handle_login_verification(driver):
            print("[ERROR] 보안 인증 페이지 감지됨")
            driver.save_screenshot("security_challenge.png")
            driver.quit()
            sys.exit(1)

        # URL 가져오기
        url = get_analytics_url()
        if not url:
            print("[ERROR] Analytics URL을 가져오지 못함")
            driver.quit()
            sys.exit(1)
        print("[INFO] Analytics URL:", url)

        # 페이지 로드
        driver.get(url)
        print("[INFO] 페이지 로드 시작...")
        
        # Analytics 페이지 로드 대기 (향상된 대기 로직)
        if not wait_for_analytics_page(driver, timeout=90):
            print("[ERROR] Analytics 페이지 로드 실패")
            driver.save_screenshot("analytics_page_failed.png")
            driver.quit()
            sys.exit(1)
            
        # 기존 로직 유지
        wait_for_page_load(driver, timeout=60)
        
        # 디버깅을 위한 페이지 구조 분석
        analyze_page_structure(driver)
        
        # 다운로드 전 스크린샷
        driver.save_screenshot("screen_before_download.png")
        
        # 다운로드 실행
        if not execute_download(driver):
            print("[ERROR] 다운로드 실패")
            driver.save_screenshot("download_failed.png")
            driver.quit()
            sys.exit(1)
        
        # 파일 처리 및 시트 업데이트
        xlsx = get_latest_xlsx(dl_dir)
        if not xlsx:
            print("[ERROR] XLSX 파일을 찾지 못함")
            driver.save_screenshot("no_xlsx_found.png")
            driver.quit()
            sys.exit(1)

        print("[INFO] 파일 경로:", xlsx)
        exposure, reached, reactions, comments, reposts, post_time = parse_excel(xlsx)
        row = get_next_row_index()
        write_metrics_to_sheet(exposure, reached, reactions, comments, reposts, row)
        write_post_time_to_sheet(post_time)
        print(f"[INFO] 시트 기록 완료 (행 {row})")

        # 임시 파일 정리
        try:
            os.remove(xlsx)
        except Exception as e:
            print("임시 파일 삭제 실패:", e)

        # 작업 완료
        driver.quit()
        print("[INFO] 작업 완료")
        
    except Exception as e:
        print(f"[ERROR] 예기치 않은 오류 발생: {e}")
        try:
            driver.save_screenshot("unexpected_error.png")
        except:
            pass
        driver.quit()
        sys.exit(1)

if __name__ == "__main__":
    main()