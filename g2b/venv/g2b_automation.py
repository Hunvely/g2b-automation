import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from datetime import datetime, timedelta
from selenium.common.exceptions import NoSuchElementException
import os
import pyhwp

import chromedriver_autoinstaller

chromedriver_autoinstaller.install()
options = Options()

# 사용자 홈 디렉토리 가져오기
home_dir = os.path.expanduser("~")  # Windows, macOS, Linux 모두 지원

# 바탕화면의 "첨부파일" 폴더 경로 설정
download_dir = os.path.join(home_dir, "Desktop", "첨부파일")

# 폴더가 없으면 생성
if not os.path.exists(download_dir):
    os.makedirs(download_dir)

print("다운로드 경로:", download_dir)

# ChromeOptions 설정
options = Options()
options.add_experimental_option("prefs", {
    "download.default_directory": download_dir,  # 다운로드 경로
    "download.prompt_for_download": False,  # 다운로드 시 사용자 확인창 표시하지 않음
    "download.directory_upgrade": True,  # 기존 경로를 업데이트
    "safebrowsing.enabled": True  # 안전 브라우징 기능 활성화
})

# 크롬 창 안 닫히게 유지
options.add_experimental_option("detach", True)

driver = webdriver.Chrome(options)

# 나라장터 페이지로 이동동
driver.get("https://www.g2b.go.kr")

# 나라장터 페이지 로드 완료될 때까지 sleep 주기
time.sleep(10)

# 창 최대화
driver.maximize_window()

def random_wait():
    time.sleep(random.uniform(1.5, 4.0))  # 1.5초에서 4초 사이의 랜덤 대기

# 팝업 닫기 함수 (조건부 처리 추가)
def close_popup(css_selector):
    try:
        # 요소가 존재하고 클릭 가능할 경우 닫기
        element = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, css_selector))
        )
        element.click()
    except TimeoutException:
        print("팝업이 없습니다.")

def scroll_until_element_visible(driver, xpath, max_scrolls=20, scroll_step=300, wait_time=1):
    for scroll_count in range(max_scrolls):
        try:
            # 첨부파일 요소가 화면에 나타났는지 확인
            element = driver.find_element(By.XPATH, xpath)
            if element.is_displayed():
                print(f"첨부파일 요소가 화면에 표시되었습니다: {xpath}")
                return True
        except NoSuchElementException:
            pass

        # 요소가 보이지 않으면 스크롤
        driver.execute_script(f"window.scrollBy(0, {scroll_step});")
        time.sleep(wait_time)

    print(f"최대 {max_scrolls}번 스크롤했지만 요소를 찾을 수 없습니다: {xpath}")
    return False

# ZIP 파일을 지정된 폴더로 추출
def extract_zip(zip_file_path, extract_to_folder):
    with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
        zip_ref.extractall(extract_to_folder)
    print(f"ZIP 파일 {zip_file_path}이(가) {extract_to_folder}에 추출되었습니다.")

# 다운로드 폴더에서 가장 최근에 다운로드된 파일을 반환
def get_latest_downloaded_file(download_dir):
    files = os.listdir(download_dir)
    files_with_path = [os.path.join(download_dir, file) for file in files]
    latest_file = max(files_with_path, key=os.path.getmtime)
    return latest_file

# Windows에서 파일을 여는 함수
def open_file(file_path):
    os.startfile(file_path)

# 다운로드된 파일의 확장자에 따라 처리
def handle_file(file_path):
    file_extension = file_path.lower().split('.')[-1]

    if file_extension == 'zip':
        # ZIP 파일 처리
        extract_folder = os.path.join(download_dir, "extracted_files")
        if not os.path.exists(extract_folder):
            os.makedirs(extract_folder)
        extract_zip(file_path, extract_folder)

    elif file_extension == 'hwpx':
        # HWPX 파일 처리
        print("HWPX 파일 처리 필요: ", file_path)

    else:
        # 기타 파일 열기
        open_file(file_path)

# 팝업 닫기 호출 (조건부 처리)
popups = [
    "#mf_wfm_container_wq_uuid_869_wq_uuid_876_poupR23AB0000013455_close",
    "#mf_wfm_container_wq_uuid_869_wq_uuid_876_poupR23AB0000013415_close",
    "#mf_wfm_container_wq_uuid_869_wq_uuid_876_poupR23AB0000013414_close",
]

for popup_selector in popups:
    close_popup(popup_selector)
    time.sleep(1)

# 발주 메뉴 클릭
ordering = "#mf_wfm_gnb_wfm_gnbMenu_wq_uuid_522"
ordering_click = driver.find_element(By.CSS_SELECTOR, ordering)
ordering_click.click()
time.sleep(3)

# 발주목록 소메뉴 클릭
ordering_list = "#mf_wfm_gnb_wfm_gnbMenu_genDepth1_0_genDepth2_0_btn_menuLvl2"
ordering_list_click = driver.find_element(By.CSS_SELECTOR, ordering_list)
ordering_list_click.click()
time.sleep(5)

# 사전규격공개 옵션 선택
pre_specification = "#mf_wfm_container_radSrchTy > li.w2radio_item.w2radio_item_1"
pre_specification_click = driver.find_element(By.CSS_SELECTOR, pre_specification)
pre_specification_click.click()
time.sleep(2)


# 어제 날짜 계산
yesterday = datetime.now() - timedelta(days=1)
yesterday_str = yesterday.strftime("%Y%m%d")

# 진행일자 시작일 input박스 클릭
start_date_xpath = "//input[@type='text' and contains(@id, 'ibxStrDay')]"
start_date_click = driver.find_element(By.XPATH, start_date_xpath)
start_date_click.click()
time.sleep(1)

# 진행일자 시작일 기존 값 제거
start_date_click.clear()
time.sleep(1)

# 진행일자 시작일 입력
start_date_click.send_keys(yesterday_str)
time.sleep(1)

# 진행일자 종료일 input박스 클릭
end_date_xpath = "//input[@type='text' and contains(@id, 'ibxEndDay')]"
end_date_click = driver.find_element(By.XPATH, end_date_xpath)
end_date_click.click()
time.sleep(1)

# 진행일자 종료일 기존 값 제거
end_date_click.clear()
time.sleep(1)

# 진행일자 종료일 입력
end_date_click.send_keys(yesterday_str)
time.sleep(1)

# 상세 조건 펼치기
detail = "#wq_uuid_1918_btnSearchToggle"
detail_click = driver.find_element(By.CSS_SELECTOR, detail)
detail_click.click()

# 업무구분 일반용역 클릭
work1 = "#mf_wfm_container_chkRqdcBsneSeCd_input_2"
work1_click = driver.find_element(By.CSS_SELECTOR, work1)
work1_click.click()
time.sleep(1)

# 업무구분 기술영역 클릭
work2 = "#mf_wfm_container_chkRqdcBsneSeCd_input_3"
work2_click = driver.find_element(By.CSS_SELECTOR, work2)
work2_click.click()
time.sleep(1)

# 사업명 입력 박스 클릭
search_box = "#mf_wfm_container_txtBizNm"
search_box_click = driver.find_element(By.CSS_SELECTOR, search_box)
search_box_click.click()
time.sleep(1)

# 사업명 입력
search_word = '구축'
search_box_click.send_keys(search_word)
time.sleep(1)

# 검색 버튼 클릭 (리스트 표시까지 완료)
search_button = "#mf_wfm_container_btnS0001"
search_button_click = driver.find_element(By.CSS_SELECTOR, search_button)
search_button_click.click()

# 리스트에 항목 있는지 확인
tbody_id = "mf_wfm_container_gridView1_body_tbody"
rows = driver.find_elements(By.CSS_SELECTOR, f"#{tbody_id} tr")
time.sleep(2)

if rows:  # tr 요소가 하나 이상 있을 경우
    for row in rows:
        # 각 row에서 링크를 찾기
        link = row.find_element(By.CSS_SELECTOR, "td a")
        
        # 링크가 클릭 가능할 때까지 대기
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable(link))
        
        # 링크 클릭
        link.click()

        # 새 페이지 로드 대기 (예: 페이지 헤더 제목이 로드될 때까지)
        try:
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "#mf_wfm_cntsHeader_spnHeaderTitle"))
            )
            print("새 페이지 로드 완료")
        except TimeoutException:
            print("새 페이지 로드 실패")
            continue  # 새 페이지 로드가 실패한 경우 다음 항목으로 넘어감

        time.sleep(2)

        # 스크롤 조건: 첨부파일이 화면에 보일 때까지
        target_xpath = "//*[@id='wq_uuid_2207_groupTitle']"
        if scroll_until_element_visible(driver, target_xpath):
            print("스크롤 완료. 첨부파일을 화면에 표시.")
        else:
            print("첨부파일을 찾지 못했습니다.")

        # 전체 선택 체크박스 클릭
        checkbox = driver.find_element(By.XPATH, "//*[contains(@id, '_header__column1_checkboxLabel__id')]")
        checkbox.click()
        time.sleep(1)

        # 파일 다운로드
        download_button = driver.find_element(By.XPATH, "//*[contains(@id, '_btnFileDown')]")
        download_button.click()
        time.sleep(1)

        # 다운로드된 최신 파일 찾기
        latest_file = get_latest_downloaded_file(download_dir)

        # 파일 열기
        if latest_file:
            print(f"최근 다운로드된 파일: {latest_file}")
            handle_file(latest_file)
        else:
            print("다운로드된 파일이 없습니다.")

else:
    search_box_click.click()
    search_word = '레포트'
    search_box_click.send_keys(search_word)
    time.sleep(1)
    search_button_click.click()

time.sleep(10)
driver.quit()

input()


