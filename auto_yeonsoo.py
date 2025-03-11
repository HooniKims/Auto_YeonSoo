import selenium
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.chrome.service import Service
import getpass
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select

from selenium.webdriver.support import expected_conditions as EC

neti_id = input("neti ID : ")
neti_pass = input("neti Password : ")

# 웹 주소 자동 설정
home_url = "https://www.neti.go.kr"
print(f"잠시 후 {home_url}로 이동합니다...")
time.sleep(2)  # 사용자가 메시지를 볼 수 있도록 잠시 대기

options = webdriver.ChromeOptions()
options.add_argument('window-size=1920,1080')
# SSL 인증서 오류 해결을 위한 추가 옵션들
options.add_argument('--ignore-certificate-errors')
options.add_argument('--ignore-ssl-errors')
options.add_argument('--allow-insecure-localhost')
options.add_argument('--disable-web-security')
options.add_argument('--allow-running-insecure-content')
options.add_experimental_option('excludeSwitches', ['enable-logging'])  # 로그 메시지 제외
options.set_capability("acceptInsecureCerts", True)  # 안전하지 않은 인증서 수락
# options.add_argument("headless")
options.add_argument("--mute-audio")
user = getpass.getuser()
user_data = r"C:\User Data"
options.add_argument(f"user-data-dir={user_data}")
options.add_argument("profile-directory=Default")
driver = webdriver.Chrome(options=options)
time.sleep(3)
driver.get(home_url)

# 팝업 창 처리
time.sleep(2)  # 팝업이 뜰 때까지 잠시 대기
main_window = driver.current_window_handle  # 메인 창 핸들 저장

# 모든 창 핸들 가져오기
handles = driver.window_handles
for handle in handles:
    if handle != main_window:
        # 팝업 창으로 전환
        driver.switch_to.window(handle)
        # popupid가 URL에 포함된 창만 닫기
        if 'popupid' in driver.current_url.lower():
            driver.close()
        
# 메인 창으로 다시 전환
driver.switch_to.window(main_window)

# 페이지가 완전히 로드될 때까지 충분히 기다림
time.sleep(5)

# 로그인 버튼 클릭 (링크 텍스트 방식)
wait = WebDriverWait(driver, 20)
login_slt = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "로그인")))
login_slt.click()

# ID 입력 (명시적 대기 추가)
time.sleep(3)  # 페이지 전환 대기 시간
id_s = '#userInputId'
id_slt = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, id_s)))
id_slt.send_keys(neti_id)

# 비밀번호 입력
pass_s = '#userInputPw'
pass_slt = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, pass_s)))
pass_slt.send_keys(neti_pass)

# 로그인 버튼 클릭 (클래스 이용)
login_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.btn_basic_color.btn_basic.one.mt30")))
login_btn.click()

time.sleep(3)

sugang3_s = '#mainLoling_99 > li > div.clearfix.card_bottom > a.btn.btn_continue'
sugang3_slt = driver.find_element(By.CSS_SELECTOR, sugang3_s)
sugang3_slt.click()

# 새로 열린 나이스 창으로 이동
driver.switch_to.window(driver.window_handles[-1])

time.sleep(3)
sugang4_s = '#lx-player > div.player_box > button.vjs-big-play-button'
sugang4_slt=driver.find_element(By.CSS_SELECTOR, sugang4_s)
sugang4_slt.click()

selector = '#next-btn'

for i in range(10000):
    try:
        # 퀴즈 완료 버튼 클릭 (클래스 이용)
        quiz_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "div.quizShowBtn.draggable")))
        quiz_btn.click()
        print("퀴즈를 마친후 클릭")
    except:
        try:
            # 다음 버튼 클릭
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
                exit(0)  # 즉시 프로그램 종료
            print("continue")
            continue

# 프로그램 정상 종료
driver.quit()