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

# 전역 변수
SPEED_ALREADY_SET = False
PLAY_BUTTON_CLICKED = False  # 재생 버튼 클릭 여부를 추적하는 변수 추가

# 사용자 입력
neti_id = input("neti ID : ")
neti_pass = input("neti Password : ")

# 웹 주소 자동 설정
home_url = "https://www.neti.go.kr"
print(f"잠시 후 {home_url}로 이동합니다...")
time.sleep(2)

# Chrome 옵션 설정
options = webdriver.ChromeOptions()
options.add_argument('--no-sandbox')
options.add_argument('--disable-dev-shm-usage')
options.add_argument('--disable-gpu')
options.add_argument('--disable-extensions')
options.add_argument('--disable-software-rasterizer')
options.add_argument('--ignore-certificate-errors')
options.add_argument('--log-level=3')
options.add_experimental_option('excludeSwitches', ['enable-logging'])

# 기본 설정 변경
prefs = {
    "profile.default_content_settings.popups": 0,
    "profile.default_content_setting_values.notifications": 2,
    "profile.default_content_setting_values.automatic_downloads": 1,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
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
    global SPEED_ALREADY_SET
    
    if SPEED_ALREADY_SET:
        return True
        
    try:
        time.sleep(2)
        
        # 배속 버튼 찾기
        speed_btn = wait.until(EC.presence_of_element_located((
            By.CSS_SELECTOR, 
            'button.vjs-playback-rate.vjs-menu-button.vjs-menu-button-popup.vjs-button[title="재생 배속"]'
        )))
        
        # 마우스 오버 동작 시뮬레이션
        action = webdriver.ActionChains(driver)
        action.move_to_element(speed_btn).perform()
        time.sleep(1)  # 메뉴가 나타날 때까지 1초 대기

        # 1.5배속 메뉴 아이템 찾기 및 클릭
        speed_1_5 = wait.until(EC.element_to_be_clickable((
            By.CSS_SELECTOR, 
            'span.vjs-menu-item-text'
        )))
        
        if speed_1_5.text == '1.5x':
            # JavaScript로 클릭 실행
            driver.execute_script("arguments[0].click();", speed_1_5)
            print("1.5배속 설정 완료")
            SPEED_ALREADY_SET = True
            return True
        else:
            print("1.5배속 옵션을 찾을 수 없음")
            return False
            
    except Exception as e:
        print(f"1.5배속 설정 중 오류 발생: {e}")
        return False

# 팝업 창 처리
def handle_popups(driver, wait):
    """팝업 창을 처리하는 함수"""
    try:
        time.sleep(2)
        main_window = driver.current_window_handle
        handles = driver.window_handles
        
        for handle in handles:
            if handle != main_window:
                driver.switch_to.window(handle)
                current_url = driver.current_url.lower()
                
                # 특정 팝업 무시 조건
                if ('selecthpgpopup.do' in current_url or 
                    'popupid=3000000835' in current_url or
                    'popupid' in current_url):
                    try:
                        # 팝업의 "오늘 하루 보지 않기" 체크박스가 있다면 클릭
                        checkbox = driver.find_element(By.CSS_SELECTOR, "input[type='checkbox']")
                        if not checkbox.is_selected():
                            checkbox.click()
                            time.sleep(0.5)
                    except:
                        pass
                    
                    try:
                        # 닫기 버튼 클릭
                        close_btn = driver.find_element(By.CSS_SELECTOR, "button.btn_close, a.btn_close, #closeBtn, .close")
                        close_btn.click()
                    except:
                        driver.close()
                        
        driver.switch_to.window(main_window)
    except Exception as e:
        print(f"팝업 처리 중 오류 발생: {str(e)}")
        driver.switch_to.window(main_window)

# 드라이버 초기화 및 웹사이트 접속
try:
    print("Chrome 브라우저를 시작합니다...")
    driver = webdriver.Chrome(options=options)
    wait = WebDriverWait(driver, 20)
    driver.get(home_url)
    print(f"{home_url}로 이동 완료")
    
    # 팝업 처리
    handle_popups(driver, wait)
    
    time.sleep(5)
except Exception as e:
    print(f"브라우저 초기화 실패: {str(e)}")
    print("\n프로그램을 종료합니다.")
    print("1. Chrome 브라우저가 설치되어 있는지 확인해주세요.")
    print("2. 인터넷 연결을 확인해주세요.")
    input("\n아무 키나 누르면 종료됩니다...")
    sys.exit(1)

# 로그인 프로세스
try:
    login_slt = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "로그인")))
    login_slt.click()
    
    # 로그인 페이지에서도 팝업 처리
    handle_popups(driver, wait)
    
    id_slt = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, '#userInputId')))
    id_slt.send_keys(neti_id)

    pass_slt = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, '#userInputPw')))
    pass_slt.send_keys(neti_pass)

    login_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.btn_basic_color.btn_basic.one.mt30")))
    login_btn.click()
    
    # 로그인 후에도 팝업 처리
    handle_popups(driver, wait)
    
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
    time.sleep(2)  # 창 전환 대기 시간 단축

    # 재생 버튼 클릭과 1.5배속 설정
    try:
        if not PLAY_BUTTON_CLICKED:  # 최초 1회만 재생 버튼 클릭
            # 재생 버튼이 보이면 클릭
            play_btn = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, 'vjs-big-play-button')))
            play_btn.click()
            PLAY_BUTTON_CLICKED = True
            time.sleep(2)  # 영상 시작 대기 시간 단축
        
        # 최초 실행 시 1회만 배속 설정 시도
        if not SPEED_ALREADY_SET:
            set_playback_speed(driver, wait)

    except Exception as e:
        print("강의 시작 과정에서 오류 발생:", str(e))

except Exception as e:
    print("강의 시작 과정에서 오류 발생:", str(e))

# 메인 루프
last_message_time = 0
message_interval = 5  # 5초 간격으로 메시지 출력
video_ended = False

while True:
    try:
        current_time = time.time()
        
        # 창이 닫혔는지 먼저 확인
        try:
            driver.current_url
        except:
            print("\n수강이 종료되었습니다.")
            driver.quit()
            
            # 카운트다운 추가
            for i in range(3, 0, -1):
                print(f"\r프로그램이 {i}초 후에 종료됩니다...", end='', flush=True)
                time.sleep(1)
            print("\n프로그램을 종료합니다.")
            sys.exit(0)
            
        if current_time - last_message_time >= message_interval:
            print("continue")
            last_message_time = current_time

        # 퀴즈 완료 버튼이 있는지 확인
        try:
            quiz_btn = driver.find_element(By.CSS_SELECTOR, "div.quizShowBtn.draggable")
            if quiz_btn.is_displayed() and quiz_btn.is_enabled():
                driver.execute_script("arguments[0].click();", quiz_btn)
                print("퀴즈 완료")
                time.sleep(0.1)  # 퀴즈 완료 후 최소 대기
                
                # 다음 버튼 즉시 클릭 시도
                try:
                    next_btn = driver.find_element(By.CSS_SELECTOR, "#next-btn")
                    if next_btn.is_displayed() and next_btn.is_enabled():
                        driver.execute_script("arguments[0].click();", next_btn)
                        print("다음 강의로 이동")
                        time.sleep(0.1)  # 페이지 전환 최소 대기
                except:
                    pass
                continue  # 퀴즈 처리 후 다음 반복으로
        except:
            pass

        # 영상 진행 상태 확인
        try:
            video_player = driver.find_element(By.CSS_SELECTOR, 'video.vjs-tech')
            current_time = driver.execute_script("return arguments[0].currentTime", video_player)
            duration = driver.execute_script("return arguments[0].duration", video_player)
            
            # 영상이 끝나면 다음으로 진행
            if not video_ended and duration > 0 and current_time >= duration - 0.5:
                video_ended = True
                print("영상 시청 완료")
                
                # 다음 버튼 즉시 클릭 시도
                try:
                    next_btn = driver.find_element(By.CSS_SELECTOR, "#next-btn")
                    if next_btn.is_displayed() and next_btn.is_enabled():
                        driver.execute_script("arguments[0].click();", next_btn)
                        print("다음 강의로 이동")
                        video_ended = False
                        time.sleep(0.1)  # 페이지 전환 최소 대기
                except:
                    pass

        except Exception as e:
            pass

        time.sleep(0.1)

    except Exception as e:
        continue