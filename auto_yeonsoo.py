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
SPEED_ALREADY_SET = False  # 1.5배속 설정 여부
PLAY_BUTTON_CLICKED = False  # 재생 버튼 클릭 여부

# 사용자 입력 받기
print("ⓒ 2025 HooniKim All rights reserved.")
neti_id = input("neti ID : ")
neti_pass = input("neti Password : ")
course_selection = input("강좌를 선택하세요 예 '1 or 2 or 3'=>")

# 메인 URL 설정
home_url = "https://www.neti.go.kr"
print(f"잠시 후 {home_url}로 이동합니다...")
time.sleep(2)

# Chrome 브라우저 설정
options = webdriver.ChromeOptions()
options.add_argument('--no-sandbox')  # 보안 설정
options.add_argument('--disable-dev-shm-usage')  # 메모리 사용량 제한
options.add_argument('--disable-gpu')  # GPU 가속 비활성화
options.add_argument('--disable-extensions')  # 확장 프로그램 비활성화
options.add_argument('--disable-software-rasterizer')  # 소프트웨어 래스터라이저 비활성화
options.add_argument('--ignore-certificate-errors')  # 인증서 오류 무시
options.add_argument('--log-level=3')  # 로그 레벨 최소화
options.add_argument('--mute-audio')  # 오디오 음소거
options.add_experimental_option('excludeSwitches', ['enable-logging'])  # 로깅 비활성화

# 브라우저 기본 설정
prefs = {
    "profile.default_content_settings.popups": 0,  # 팝업 차단
    "profile.default_content_setting_values.notifications": 2,  # 알림 차단
    "profile.default_content_setting_values.automatic_downloads": 1,  # 자동 다운로드 허용
    "download.prompt_for_download": False,  # 다운로드 프롬프트 비활성화
    "download.directory_upgrade": True,  # 다운로드 디렉토리 업그레이드
    "safebrowsing.enabled": True  # 안전 브라우징 활성화
}
options.add_experimental_option("prefs", prefs)

# 사용자 프로필 설정
try:
    user = getpass.getuser()
    user_data = os.path.join(os.path.dirname(os.path.abspath(__file__)), "User Data")
    if not os.path.exists(user_data):
        os.makedirs(user_data)
    options.add_argument(f"user-data-dir={user_data}")
    options.add_argument("profile-directory=Default")
except Exception as e:
    print(f"사용자 데이터 설정 중 오류 발생: {str(e)}")

# 재생 속도 설정 함수
def set_playback_speed(driver, wait):
    """영상 재생 속도를 1.5배속으로 설정"""
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
        
        # 배속 메뉴 표시
        action = webdriver.ActionChains(driver)
        action.move_to_element(speed_btn).perform()
        time.sleep(1)

        # 1.5배속 옵션 선택
        speed_1_5 = wait.until(EC.element_to_be_clickable((
            By.CSS_SELECTOR, 
            'span.vjs-menu-item-text'
        )))
        
        if speed_1_5.text == '1.5x':
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

# 팝업 처리 함수
def handle_popups(driver, wait):
    """불필요한 팝업 창들을 자동으로 처리"""
    try:
        time.sleep(2)
        main_window = driver.current_window_handle
        handles = driver.window_handles
        
        for handle in handles:
            if handle != main_window:
                driver.switch_to.window(handle)
                current_url = driver.current_url.lower()
                
                # 무시할 팝업 URL 확인
                if ('selecthpgpopup.do' in current_url or 
                    'popupid=3000000835' in current_url or
                    'popupid' in current_url):
                    try:
                        # "오늘 하루 보지 않기" 체크
                        checkbox = driver.find_element(By.CSS_SELECTOR, "input[type='checkbox']")
                        if not checkbox.is_selected():
                            checkbox.click()
                            time.sleep(0.5)
                    except:
                        pass
                    
                    try:
                        # 팝업 닫기
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

# 강의 선택 및 시작
try:
    # 나의 학습방 메뉴 클릭
    my_study_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#container > div.main_top_bg.clearfix > div.por.pagew > div.service_type1 > div.frequent_service_wrap.clearfix.frequent_aa > div.frequent_service_list.clearfix.frequent_bb > ul > li:nth-child(2) > a')))
    my_study_btn.click()
    time.sleep(5)

    # 강좌 선택
    try:
        if course_selection == "1":
            course_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#crseList > li:nth-child(1) > div.txt_box > div.btn_box.mt0 > a.bnt_basic_line.small')))
            course_btn.click()
        elif course_selection == "2":
            course_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#crseList > li:nth-child(2) > div.txt_box > div.btn_box.mt0 > a.bnt_basic_line.small')))
            course_btn.click()
        elif course_selection == "3":
            course_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#crseList > li:nth-child(3) > div.txt_box > div.btn_box.mt0 > a.bnt_basic_line.small')))
            course_btn.click()
    except Exception as e:
        print(f"강좌 선택 중 오류 발생: {str(e)}")
        driver.quit()
        exit(1)

    time.sleep(2)

    # 학습 콘텐츠 버튼이 있는지 확인하고 클릭
    try:
        # JavaScript로 학습 콘텐츠 버튼 클릭
        driver.execute_script("""
            var element = document.querySelector('#conList > ul.learning_list.listMain.remote > li > div > ul > li.clearfix.on > div > div:nth-child(1) > div > div > div > div.learn_con > a');
            if (element) {
                element.click();
            }
        """)
        print("학습 콘텐츠 버튼 클릭 완료")
        print("처음 영상 재생은 조금 시간이 걸릴 수 있습니다. 기다려 주세요.\n다음 영상으로 넘어갈 때는 느리지 않습니다.")
            
    except Exception as e:
        print(f"학습 콘텐츠 버튼 클릭 중 오류 발생: {str(e)}")
        print("현재 선택된 강좌로 진행합니다.")
        pass

    # 새 창으로 전환
    driver.switch_to.window(driver.window_handles[-1])

    # 강의 시작 버튼 클릭
    try:
        lec_button_xpath = '//*[@id="conList"]/ul[1]/li/div/ul/li[1]/div/div[1]/div/div/div/div[2]/a'
        lec_survey_button = wait.until(EC.element_to_be_clickable((By.XPATH, lec_button_xpath)))
        lec_survey_button.click()
    except:
        pass

    # 강의 보기 버튼 클릭
    try:
        btn = driver.find_element(By.CSS_SELECTOR, '#lectBtnControl > p').click()
    except:
        pass

    # 플레이어 컨트롤
    try:
        # 플레이 버튼 클릭
        btn = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, 'vjs-big-play-button')))
        btn.click()
        
        # 1.5배속 설정
        if not SPEED_ALREADY_SET:
            try:
                # 배속 버튼 찾기
                speed_btn = wait.until(EC.presence_of_element_located((
                    By.CSS_SELECTOR, 
                    'button.vjs-playback-rate.vjs-menu-button.vjs-menu-button-popup.vjs-button[title="재생 배속"]'
                )))
                
                # 마우스 오버로 배속 메뉴 표시
                action = webdriver.ActionChains(driver)
                action.move_to_element(speed_btn).perform()

                # 1.5배속 옵션 선택
                speed_1_5 = wait.until(EC.element_to_be_clickable((
                    By.CSS_SELECTOR, 
                    'span.vjs-menu-item-text'
                )))
                
                if speed_1_5.text == '1.5x':
                    driver.execute_script("arguments[0].click();", speed_1_5)
                    print("1.5배속 설정 완료")
                    SPEED_ALREADY_SET = True
            except Exception as e:
                print(f"1.5배속 설정 중 오류 발생: {e}")
    except:
        pass

except Exception as e:
    print("강의 시작 과정에서 오류 발생:", str(e))
    driver.quit()
    exit(1)

# 메인 루프
last_message_time = 0
message_interval = 60  # 60초(1분) 간격으로 메시지 출력
video_ended = False
last_progress = 0

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

        # 마지막 목차인지 확인
        try:
            last_chapter = driver.find_element(By.CSS_SELECTOR, 'p.desc[style="line-height:150%"][tabindex="-1"]')
            if last_chapter.text == "마지막 목차입니다.":
                print("\n모든 강의가 종료되었습니다.")
                print("\n아무 키나 눌러주세요.")
                input()
                print("\n프로그램을 종료합니다.")
                driver.quit()
                sys.exit(0)
        except:
            pass

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
            
            # 영상이 실제로 재생 중인지 확인
            if current_time > last_progress:
                last_progress = current_time
            
            # 영상이 끝나면 다음으로 진행
            if not video_ended and duration > 0 and current_time >= duration - 0.1:
                # 재생이 실제로 완료되었는지 한 번 더 확인
                time.sleep(0.5)  # 잠시 대기 후 재확인
                current_time = driver.execute_script("return arguments[0].currentTime", video_player)
                
                if current_time >= duration - 0.1:
                    video_ended = True
                    print("영상 시청 완료")
                    
                    # 다음 버튼 즉시 클릭 시도
                    try:
                        next_btn = driver.find_element(By.CSS_SELECTOR, "#next-btn")
                        if next_btn.is_displayed() and next_btn.is_enabled():
                            driver.execute_script("arguments[0].click();", next_btn)
                            print("다음 강의로 이동")
                            video_ended = False
                            last_progress = 0  # 진행 상태 초기화
                            time.sleep(0.1)  # 페이지 전환 최소 대기
                    except:
                        pass

        except Exception as e:
            pass

        time.sleep(0.1)

    except Exception as e:
        continue