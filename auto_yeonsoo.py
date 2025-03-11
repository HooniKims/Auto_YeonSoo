import selenium
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.support.select import Select
import time
import getpass
import os
import sys
from subprocess import CREATE_NO_WINDOW

# 사용자 입력
neti_id = input("neti ID : ")
neti_pass = input("neti Password : ")

# 웹 주소 자동 설정
home_url = "https://www.neti.go.kr"
print(f"잠시 후 {home_url}로 이동합니다...")
time.sleep(2)

# Chrome 옵션 설정
options = webdriver.ChromeOptions()
options.add_argument('window-size=1920,1080')

# SSL/보안 관련 옵션 강화
options.add_argument('--ignore-certificate-errors')
options.add_argument('--ignore-ssl-errors')
options.add_argument('--ignore-certificate-errors-spki-list')
options.add_argument('--allow-insecure-localhost')
options.add_argument('--disable-web-security')
options.add_argument('--allow-running-insecure-content')
options.add_argument('--reduce-security-for-testing')
options.add_argument('--no-sandbox')
options.add_argument('--disable-dev-shm-usage')

# 안정성 향상을 위한 옵션들
options.add_argument('--disable-gpu')
options.add_argument('--disable-features=VizDisplayCompositor')
options.add_argument('--disable-blink-features=AutomationControlled')
options.add_experimental_option('excludeSwitches', ['enable-logging', 'enable-automation'])
options.add_experimental_option('useAutomationExtension', False)
options.set_capability("acceptInsecureCerts", True)
options.add_argument("--mute-audio")

# 프록시 설정 비활성화 (선택적)
options.add_argument('--no-proxy-server')

# 기본 설정 변경
prefs = {
    "profile.default_content_settings.popups": 0,
    "profile.default_content_setting_values.notifications": 2,
    "profile.default_content_setting_values.automatic_downloads": 1,
    "profile.default_content_settings.geolocation": 2,
    "profile.managed_default_content_settings.images": 1,
    "profile.default_content_setting_values.mixed_script": 1,
    "profile.content_settings.exceptions.automatic_downloads.*.setting": 1,
    "profile.default_content_setting_values.insecure_content": "allow"
}
options.add_experimental_option("prefs", prefs)

# 사용자 데이터 설정
try:
    user = getpass.getuser()
    user_data = os.path.join(os.path.dirname(os.path.abspath(__file__)), "User Data")
    if not os.path.exists(user_data):
        os.makedirs(user_data)
    options.add_argument(f"user-data-dir={user_data}")
    options.add_argument("profile-directory=Default")
except Exception as e:
    print(f"사용자 데이터 설정 중 오류 발생: {str(e)}")

# 함수들을 코드 시작 부분에 배치
def set_playback_speed(driver, wait):
    """재생 속도를 1.5배속으로 설정하는 함수"""
    try:
        # 배속 버튼 클릭
        speed_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'button.vjs-playback-rate[title="재생 배속"]')))
        speed_btn.click()
        time.sleep(2)  # 메뉴가 완전히 열릴 때까지 대기

        # 정확한 선택자를 사용하여 1.5배속 선택
        speed_1_5 = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'li.vjs-menu-item[role="menuitemradio"] span.vjs-menu-item-text')))
        if speed_1_5.text == '1.5x':
            speed_1_5.click()
            return True
        return False
    except Exception as e:
        return False

def get_chromedriver_path():
    current_dir = os.path.dirname(os.path.abspath(__file__))
    chromedriver_name = "chromedriver.exe"
    chromedriver_path = os.path.join(current_dir, chromedriver_name)
    return chromedriver_path

def initialize_driver():
    try:
        service = Service(get_chromedriver_path())
        service.creation_flags = CREATE_NO_WINDOW
        driver = webdriver.Chrome(service=service, options=options)
        return driver
    except Exception as e:
        print(f"드라이버 초기화 실패: {str(e)}")
        return None

# 드라이버 초기화 시도
max_attempts = 3
for attempt in range(max_attempts):
    try:
        driver = initialize_driver()
        if driver:
            wait = WebDriverWait(driver, 20)
            driver.get(home_url)
            # 페이지 로드 후 잠시 대기
            time.sleep(5)
            break
        else:
            raise Exception("드라이버 초기화 실패")
    except Exception as e:
        print(f"드라이버 초기화 시도 {attempt + 1}/{max_attempts} 실패: {str(e)}")
        if attempt == max_attempts - 1:
            print("드라이버 초기화 실패. 프로그램을 종료합니다.")
            input("아무 키나 누르면 종료됩니다...")
            sys.exit(1)
        time.sleep(2)

# 팝업 창 처리
time.sleep(2)
main_window = driver.current_window_handle
handles = driver.window_handles
for handle in handles:
    if handle != main_window:
        try:
            driver.switch_to.window(handle)
            if 'popupid' in driver.current_url.lower():
                driver.close()
        except:
            continue

driver.switch_to.window(main_window)
time.sleep(3)

# 로그인 프로세스
try:
    login_slt = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "로그인")))
    login_slt.click()

    id_slt = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, '#userInputId')))
    id_slt.send_keys(neti_id)

    pass_slt = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, '#userInputPw')))
    pass_slt.send_keys(neti_pass)

    login_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.btn_basic_color.btn_basic.one.mt30")))
    login_btn.click()
except Exception as e:
    print("로그인 과정에서 오류 발생:", str(e))
    driver.quit()
    exit(1)

time.sleep(3)

# 강의 시작
try:
    sugang_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#mainLoling_99 > li > div.clearfix.card_bottom > a.btn.btn_continue')))
    sugang_btn.click()

    # 새 창으로 전환
    driver.switch_to.window(driver.window_handles[-1])
    time.sleep(3)

    # 재생 버튼 클릭과 1.5배속 설정
    try:
        play_btn = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, 'vjs-big-play-button')))
        play_btn.click()
        time.sleep(3)  # 영상이 시작될 때까지 충분히 대기
        
        # 1.5배속 설정
        set_playback_speed(driver, wait)

    except Exception as e:
        print("강의 시작 과정에서 오류 발생:", str(e))

except Exception as e:
    print("강의 시작 과정에서 오류 발생:", str(e))

# 메인 루프
last_message_time = 0
message_interval = 5  # 5초 간격으로 메시지 출력

while True:
    try:
        current_time = time.time()
        
        if current_time - last_message_time >= message_interval:
            print("continue")
            last_message_time = current_time

        # 재생 버튼 클릭 시도
        try:
            play_btn = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, 'vjs-big-play-button')))
            play_btn.click()
            time.sleep(3)
            
            # 재생 시작 후 배속 재설정
            set_playback_speed(driver, wait)
        except:
            # 창이 닫혔는지 확인
            try:
                driver.current_url
            except:
                print("\n수강이 종료되었습니다. 잠시 후 프로그램이 종료됩니다.")
                time.sleep(3)  # 3초 대기 후 종료
                driver.quit()
                exit(0)
            pass

        # 퀴즈 완료 버튼 클릭 시도
        try:
            quiz_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "div.quizShowBtn.draggable")))
            quiz_btn.click()
            print("퀴즈를 마친후 클릭")
        except:
            pass

        # 다음 버튼 클릭 시도
        try:
            next_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#next-btn")))
            next_btn.click()
            print("next버튼 클릭")
        except:
            pass

        time.sleep(1)

    except Exception as e:
        continue