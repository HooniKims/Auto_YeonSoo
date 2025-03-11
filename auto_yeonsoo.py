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
# 안정성 향상을 위한 추가 옵션들
options.add_argument('--no-sandbox')
options.add_argument('--disable-dev-shm-usage')
options.add_argument('--disable-gpu')
options.add_argument('--disable-features=VizDisplayCompositor')
options.add_argument('--disable-blink-features=AutomationControlled')
# SSL 인증서 오류 해결을 위한 옵션들
options.add_argument('--ignore-certificate-errors')
options.add_argument('--ignore-ssl-errors')
options.add_argument('--allow-insecure-localhost')
options.add_argument('--disable-web-security')
options.add_argument('--allow-running-insecure-content')
options.add_experimental_option('excludeSwitches', ['enable-logging'])
options.add_experimental_option('useAutomationExtension', False)
options.set_capability("acceptInsecureCerts", True)
options.add_argument("--mute-audio")

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

# 드라이버 초기화 (재시도 로직 추가)
max_attempts = 3
for attempt in range(max_attempts):
    try:
        driver = webdriver.Chrome(options=options)
        wait = WebDriverWait(driver, 20)
        driver.get(home_url)
        break
    except Exception as e:
        print(f"드라이버 초기화 시도 {attempt + 1}/{max_attempts} 실패: {str(e)}")
        if attempt == max_attempts - 1:
            print("드라이버 초기화 실패. 프로그램을 종료합니다.")
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

        # 1.5배속 설정
        try:
            # 배속 버튼 클릭 (수정된 셀렉터)
            speed_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'button.vjs-playback-rate[title="재생 배속"]')))
            speed_btn.click()
            time.sleep(2)
            # 1.5배속 선택
            speed_1_5 = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(@class, 'vjs-menu-item-text') and text()='1.5x']")))
            speed_1_5.click()
        except Exception as e:
            print("재생 속도 설정 중 오류 발생:", str(e))

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
        
        # 상태 메시지 출력 (5초 간격)
        if current_time - last_message_time >= message_interval:
            print("continue")
            last_message_time = current_time

        # 쿠키 갱신 (1시간마다)
        if current_time % 3600 < 1:
            for cookie in driver.get_cookies():
                driver.add_cookie(cookie)

        # 재생 버튼 클릭 시도
        try:
            play_btn = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, 'vjs-big-play-button')))
            play_btn.click()
            
            # 수정된 1.5배속 재설정
            time.sleep(1)
            speed_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'button.vjs-playback-rate[title="재생 배속"]')))
            speed_btn.click()
            time.sleep(1)
            speed_1_5 = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(@class, 'vjs-menu-item-text') and text()='1.5x']")))
            speed_1_5.click()
        except:
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
            # 창이 닫혔는지 확인
            try:
                driver.current_url
            except:
                print("\n수강이 종료되었습니다. 창을 닫으셔도 좋습니다.")
                driver.quit()
                exit(0)

        time.sleep(1)

    except Exception as e:
        print(f"예상치 못한 오류 발생: {str(e)}")
        continue