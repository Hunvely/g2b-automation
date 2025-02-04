import sys
import time
import psutil
import logging
import random
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from datetime import datetime, timedelta
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import StaleElementReferenceException
import os
import zipfile
import chromedriver_autoinstaller
import pyautogui
import pywinauto
import win32com.client
import pandas as pd
from openpyxl import load_workbook


# 로깅 설정
logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
)

chromedriver_autoinstaller.install()
options = Options()

# 사용자 홈 디렉토리 가져오기
home_dir = os.path.expanduser("~")  # Windows, macOS, Linux 모두 지원

# 한글 파일 경로 (Neo 버전)
hanword_path = r"C:\\Program Files (x86)\\Hnc\\Office NEO\\HOffice96\\Bin\\Hwp.exe"

# 워드 파일 경로
word_path = r"C:\\Program Files\\Microsoft Office\\root\\Office16\WINWORD.EXE"

# PDF 파일일 경로
acrobat_path = r"C:\Program Files\Adobe\Acrobat DC\Acrobat\\Acrobat.exe"

# 엑셀 파일 경로
excel_path = r"C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE"

# 사용자 바탕화면에 있는 "스크린샷" 폴더 경로 설정
screenshot_dir = os.path.join(home_dir, "Desktop", "스크린샷")

# 바탕화면의 "첨부파일" 폴더 경로 설정
download_dir = os.path.join(home_dir, "Desktop", "첨부파일")

# 폴더가 없으면 생성
if not os.path.exists(download_dir):
    os.makedirs(download_dir)

print("다운로드 경로:", download_dir)

# ChromeOptions 설정
options = Options()
options.add_experimental_option(
    "prefs",
    {
        "download.default_directory": download_dir,  # 다운로드 경로
        "download.prompt_for_download": False,  # 다운로드 시 사용자 확인창 표시하지 않음
        "download.directory_upgrade": True,  # 기존 경로를 업데이트
        "safebrowsing.enabled": True,  # 안전 브라우징 기능 활성화
    },
)

# 크롬 창 안 닫히게 유지
options.add_experimental_option("detach", True)

driver = webdriver.Chrome(options)


def random_wait():
    time.sleep(random.uniform(1.5, 4.0))  # 1.5초에서 4초 사이의 랜덤 대기


# 팝업 닫기 함수 (조건부 처리 추가)
def close_popup(css_selector):
    try:
        # 요소가 존재하고 클릭 가능할 경우 닫기
        element = WebDriverWait(driver, 2).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, css_selector))
        )
        element.click()
        logging.info(f"팝업 닫음: {css_selector}")
    except TimeoutException:
        logging.info(f"팝업 없음: {css_selector}")


def scroll_until_element_visible(driver, max_scrolls=5, scroll_step=300, wait_time=1):
    target_xpath = "//*[contains(@id, 'wq_uuid_') and contains(@class, 'w2textbox')]"
    for scroll_count in range(max_scrolls):
        try:
            # 첨부파일 요소가 화면에 나타났는지 확인
            element = driver.find_element(By.XPATH, target_xpath)
            if element.is_displayed():
                return True
        except NoSuchElementException:
            pass

        # 요소가 보이지 않으면 스크롤
        driver.execute_script(f"window.scrollBy(0, {scroll_step});")
        time.sleep(wait_time)

    logging.warning(f"최대 {max_scrolls}번 스크롤했지만 요소를 찾을 수 없습니다: {id}")
    return False


# 상세규격정보 페이지 데이터 추출
def extract_data(driver):

    # 각 요소 찾기
    사전규격등록번호 = driver.find_element(
        By.CSS_SELECTOR, "input[title^='사전규격등록번호']"
    )
    사전규격명 = driver.find_element(By.CSS_SELECTOR, "td[data-title^='사전규격명']")
    수요기관 = driver.find_element(By.CSS_SELECTOR, "input[title^='수요기관']")
    공고기관 = driver.find_element(By.CSS_SELECTOR, "input[title^='공고기관']")
    담당자 = driver.find_element(
        By.CSS_SELECTOR, "input[title^='공고기관담당자명(전화번호)']"
    )
    배정예산액 = driver.find_element(
        By.CSS_SELECTOR, "input[title^='배정예산액(부가세포함)']"
    )
    의견등록마감일시 = driver.find_element(By.CSS_SELECTOR, "input[title^='시분']")

    # 사전규격 상세정보 필드 확인 및 데이터 추출
    try:
        상세정보_컨테이너 = driver.find_element(By.ID, "mf_wfm_container_grpUrlInfo")
        상세정보_링크 = 상세정보_컨테이너.find_element(
            By.CSS_SELECTOR, "#mf_wfm_container_ancPbancInstUrl"
        )
        상세정보_텍스트 = 상세정보_컨테이너.text.strip() or "N/A"
        # JavaScript URL 대신 텍스트에서 URL 추출
        if 상세정보_링크.get_attribute("href") == "javascript:void(null);":
            상세정보_URL = 상세정보_링크.text.strip() or "N/A"
        else:
            상세정보_URL = 상세정보_링크.get_attribute("href") or "N/A"
    except Exception:
        상세정보_텍스트 = "N/A"
        상세정보_URL = "N/A"

    try:
        # 데이터 추출
        data = {
            "사전규격등록번호": 사전규격등록번호.get_attribute("value") or "N/A",
            "사전규격명": 사전규격명.text or "N/A",
            "수요기관": 수요기관.get_attribute("value") or "N/A",
            "공고기관": 공고기관.get_attribute("value") or "N/A",
            "담당자": 담당자.get_attribute("value") or "N/A",
            "배정예산액": 배정예산액.get_attribute("value") or "N/A",
            "의견등록마감일시": 의견등록마감일시.get_attribute("value") or "N/A",
            "사전규격상세정보_URL": 상세정보_URL,
        }
        return data
    except Exception as e:
        logging.error(f"데이터 추출 중 오류 발생: {e}")
        return None


def save_to_excel(data):
    try:
        desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
        excel_folder = os.path.join(desktop_path, "엑셀파일")

        if not os.path.exists(excel_folder):
            logging.info(f"폴더가 존재하지 않음. 폴더를 생성합니다: {excel_folder}")
            os.makedirs(excel_folder)

        file_name = "사전규격상세정보.xlsx"
        full_path = os.path.join(excel_folder, file_name)

        logging.info(f"엑셀 파일 경로: {full_path}")

        sheet_name = "데이터"
        df = pd.DataFrame([data])  

        if os.path.exists(full_path):
            logging.info(f"기존 파일이 존재합니다. 데이터를 추가합니다: {full_path}")

            # 기존 데이터 불러오기
            existing_df = pd.read_excel(full_path, sheet_name=sheet_name, engine="openpyxl")

            # 새로운 데이터를 기존 데이터와 합치기
            combined_df = pd.concat([existing_df, df], ignore_index=True)

            # 합친 데이터를 다시 저장
            with pd.ExcelWriter(full_path, engine="openpyxl", mode="w") as writer:
                combined_df.to_excel(writer, index=False, sheet_name=sheet_name)
        else:
            logging.info("새로운 엑셀 파일을 생성합니다.")
            with pd.ExcelWriter(full_path, engine="openpyxl", mode="w") as writer:
                df.to_excel(writer, index=False, sheet_name=sheet_name)

        # 엑셀 파일 열기 (openpyxl로 셀 크기 조정)
        wb = load_workbook(full_path)
        sheet = wb[sheet_name]

        # 셀 크기 자동 조정 코드 개선
        for col in sheet.columns:
            max_length = 0
            column = col[0].column_letter  # 열의 알파벳(A, B, C 등)

            for cell in col:
                try:
                    if cell.value:
                        # 한글과 영어 너비 차이 보정
                        text = str(cell.value)
                        text_length = sum(2 if ord(char) > 127 else 1 for char in text)  # 한글이면 2, 영어는 1로 계산
                        max_length = max(max_length, text_length)
                except:
                    pass

            adjusted_width = (max_length * 1.2)  # 한글 보정 계수 적용
            sheet.column_dimensions[column].width = adjusted_width

        # 파일 저장
        wb.save(full_path)

        logging.info(f"데이터가 {full_path} 파일에 추가되었습니다.")
    except Exception as e:
        logging.error(f"엑셀 저장 중 오류 발생: {e}")


# 첨부파일 폴더에서 가장 최근에 다운로드된 파일을 반환
def get_latest_downloaded_file(download_dir):
    files = os.listdir(download_dir)
    files_with_path = [os.path.join(download_dir, file) for file in files]
    latest_file = max(files_with_path, key=os.path.getmtime)
    return latest_file


# Windows에서 파일을 여는 함수 (공통)
def open_file(file_path):
    os.startfile(file_path)


# ============================================== 한글 함수 ==============================================


# def close_warning_window_hangle(app):
#     try:
#         windows = app.windows()
#         for win in windows:
#             try:
#                 rect = win.rectangle()
#                 width = rect.right - rect.left
#                 height = rect.bottom - rect.top
#                 title = win.window_text()

#                 # 조건: 버전 차이 경고 메세지
#                 if width == 201 and height == 241:
#                     print(f"경고 창 감지: {title} (크기: {width}x{height})")
#                     # 창 위치로 마우스 이동 및 Enter 키 입력
#                     x, y = rect.left + 10, rect.top + 10  # 창 내부로 마우스 이동
#                     pyautogui.click(x, y)
#                     pyautogui.press("enter")
#                     time.sleep(1)
#                     return True
#             except Exception as e:
#                 print(f"창 탐색 중 오류 발생: {e}")
#         return False
#     except Exception as e:
#         print(f"경고 창을 닫는 중 오류 발생: {e}")
#         return False
    
def close_hwp_file():
    # 한글 프로세스 종료
    for proc in psutil.process_iter(["pid", "name"]):
        if "Hwp.exe" in proc.info["name"]:  # Hwp 프로세스 이름 확인
            os.kill(proc.info["pid"], 9)  # 프로세스 강제 종료
            print("Hwp가 종료되었습니다.")
            return
    print("Hwp가 실행 중이 아닙니다.")

# 다운로드된 한글 파일을 열고, 키워드를 검색하여 스크린샷을 찍는 함수 호출
def handle_hwp_file(file_path, keywords, 사전규격명):

    # 파일 존재 여부 확인
    if not os.path.exists(file_path):
        print(f"파일을 찾을 수 없습니다: {file_path}")
        return
    
    open_file(file_path)

    for keyword in keywords:  # 순차적으로 각 키워드 처리
        # 키워드 검색 후 스크린샷 찍기
        screenshot_hwp(keyword, 사전규격명)

    close_hwp_file()

def close_warning_window_hangle(app):
    try:
        alert_windows = app.windows()
        for alert in alert_windows:
            # '문서의 끝까지 찾았습니다' 텍스트가 정확히 포함된 창을 찾음
            if alert.window_text() and "문서의 끝까지 찾았습니다" in alert.window_text():
                print("검색 종료 창 감지. 검색을 종료합니다.")
                alert.set_focus()  # 창을 선택하고
                pyautogui.press("esc")  # ESC 키를 눌러 창 닫기
                return True  # 창을 찾았다면 True 반환
        return False  # 해당 창을 찾지 못했다면 False 반환
    except Exception as e:
        print(f"검색 종료 창 감지 중 오류 발생: {e}")
        return False

def screenshot_hwp(keyword, 사전규격명):
    # 한글 프로그램 자동화
    try:
        app = pywinauto.Application().connect(path=hanword_path) # 한글 프로그램 경로

        # 한글 로딩
        time.sleep(5)

        # 경고 창 닫기
        # if close_warning_window_hangle(app):
        #     print("경고 메시지가 닫혔습니다.")
        #     time.sleep(2)

        pyautogui.press("esc")  # ESC 키를 눌러 창 닫기
        time.sleep(2)

        hwp_window = app.window(title_re=".*한글.*")  # 한글 프로그램의 창을 찾기

        # 문서의 맨 위로 이동 (Ctrl + Page Up)
        hwp_window.type_keys("^({PGUP})")
        logging.info("문서 맨 위로 이동")
        time.sleep(1)

        # 모든 컨트롤 요소들 출력 (child_window)
        # hwp_window.print_control_identifiers()
        
        # 키워드 검색 (단, 한글 프로그램에서 키워드 검색 기능을 자동화하려면 단축키 활용)
        hwp_window.type_keys("^f")  # Ctrl+F (검색 단축키)
        logging.info("검색 모달 표시")
        time.sleep(2)

        # 한글 메인 편집창 찾기
        hwp_edit = app.window(title_re=".*찾기.*")
        # hwp_edit.print_control_identifiers()

        if not hwp_edit:
            print("한글 편집창을 찾을 수 없습니다.")
            return

        # 포커스를 주고 키워드 입력
        hwp_edit.set_focus()
        time.sleep(1)
        hwp_edit.type_keys(keyword, with_spaces=True, pause=0.1)
        logging.info(f"검색어 입력: {keyword}")
        time.sleep(1)

        # 엔터 키 입력 (검색 실행)
        hwp_edit.type_keys("{ENTER}")
        logging.info("검색 실행")

        # 검색된 텍스트 영역이 활성화되도록 대기
        time.sleep(2)

        capture_count = 0

        while True:

            # hwp_edit_complete 창 확인 (모든 검색 완료 후 종료)
            hwp_edit_complete = app.window(title_re="한글")
            if hwp_edit_complete.exists():
                print(f"'{keyword}'에 대한 모든 검색을 마쳤습니다.")
                pyautogui.press("esc")  # ESC 키를 눌러 창 닫기
                return True  # 모든 검색 종료

            try:
                # 스크린샷 영역 설정
                x1, y1 = 100, 200  # 좌측 상단 좌표
                width, height = 1800, 800
                x2, y2 = x1 + width, y1 + height  # 우측 하단 좌표 계산

                # 스크린샷 찍기
                screenshot = pyautogui.screenshot(region=(x1, y1, width, height))
                capture_count += 1
                screenshot_file = os.path.join(screenshot_dir, f"{사전규격명}_{keyword}_{capture_count}.png")
                screenshot.save(screenshot_file)
                print(f"검색 결과 {capture_count} 캡처 완료: {screenshot_file}")

                # "다음 찾기" 버튼 클릭
                hwp_edit.type_keys("{ENTER}")
                time.sleep(2)  # 다음 검색 결과가 로드될 시간 대기
                # hwp_edit.print_control_identifiers()

            except Exception as e:
                print(f"검색 결과 끝 또는 오류: {e}")
                break

        # ESC 키 한 번 누르기
        pyautogui.press("esc")
        print(f"총 {capture_count}개의 검색 결과 캡처 완료.")

    except Exception as e:
        print(f"한글 파일 처리 중 오류 발생: {e}")


# ============================================== hwpx 함수 ==============================================


# 다운로드된 hwpx 파일을 열고, 키워드를 검색하여 스크린샷을 찍는 함수 호출
def handle_hwpx_file(file_path, keywords, 사전규격명):
    # 파일 존재 여부 확인
    if not os.path.exists(file_path):
        print(f"파일을 찾을 수 없습니다: {file_path}")
        return

    open_file(file_path)

    for keyword in keywords:  # 순차적으로 각 키워드 처리
        # 키워드 검색 후 스크린샷 찍기
        screenshot_hwpx(keyword, 사전규격명)


def screenshot_hwpx(keyword, 사전규격명):
    screenshot_hwp(keyword, 사전규격명)  # 한글 처리와 동일 (테스트 예정)


# ============================================== PDF 함수 ==============================================


# 다운로드된 PDF 파일을 열고, 키워드를 검색하여 스크린샷을 찍는 함수 호출
def handle_pdf_file(file_path, keywords, 사전규격명):
    # 파일 존재 여부 확인
    if not os.path.exists(file_path):
        print(f"파일을 찾을 수 없습니다: {file_path}")
        return

    open_file(file_path)

    for keyword in keywords:  # 순차적으로 각 키워드 처리
        # 키워드 검색 후 스크린샷 찍기
        screenshot_pdf(keyword, 사전규격명)

    # 작업이 끝난 후 Acrobat Reader 닫기
    close_acrobat_reader()


def screenshot_pdf(keyword, 사전규격명):
    try:
        # Adobe Acrobat Reader 연결 (경로 필요 시 명시적으로 설정)
        app = pywinauto.Application().connect(path=acrobat_path)

        pdf_window = app.window(title_re=".*Adobe Acrobat Reader.*")

        # PDF 뷰어 로딩 대기
        time.sleep(7)

        # 키워드 검색 모드 활성화 (Ctrl+F)
        pdf_window.type_keys("^f")
        print("검색 모달 표시")
        time.sleep(2)

        # 키워드 입력
        pdf_window.type_keys(keyword)
        print(f"검색어 '{keyword}' 입력")
        time.sleep(1)

        # 검색 시작 (Enter)
        pdf_window.type_keys("{ENTER}")
        print("검색 시작")
        time.sleep(2)

        # 검색된 텍스트 영역이 활성화되도록 대기
        time.sleep(2)

        # 스크린샷 캡처
        capture_count = 0
        while True:
            try:
                # 화면 캡처 영역 (적절히 조정 필요)
                x1, y1 = 100, 100  # 좌측 상단 좌표
                width, height = 2000, 800  # 캡처 영역 크기
                x2, y2 = x1 + width, y1 + height

                # 스크린샷 저장
                screenshot = pyautogui.screenshot(region=(x1, y1, width, height))
                capture_count += 1
                screenshot_file = os.path.join(
                    screenshot_dir, f"{사전규격명}_{keyword}_{capture_count}.png"
                )
                screenshot.save(screenshot_file)
                print(f"검색 결과 {capture_count} 캡처 완료: {screenshot_file}")

                # 다음 검색 결과로 이동 (Enter 키)
                pdf_window.type_keys("{ENTER}")
                time.sleep(2)

                # 검색 완료 시 처리 필요

            except Exception as e:
                print(f"검색 결과 끝 또는 오류: {e}")
                break  # 루프 종료

        print(f"총 {capture_count}개의 검색 결과 캡처 완료.")

    except Exception as e:
        print(f"PDF 파일 처리 중 오류 발생: {e}")


def close_acrobat_reader():
    # Acrobat Reader 프로세스 종료
    for proc in psutil.process_iter(["pid", "name"]):
        if "Acrobat.exe" in proc.info["name"]:  # Acrobat Reader 프로세스 이름 확인
            os.kill(proc.info["pid"], 9)  # 프로세스 강제 종료
            print("Acrobat Reader가 종료되었습니다.")
            return
    print("Acrobat Reader가 실행 중이 아닙니다.")


# ============================================== 워드 함수 ==============================================


# 다운로드된 워드 파일을 열고, 키워드를 검색하여 스크린샷을 찍는 함수 호출
def handle_docx_file(file_path, keywords, 사전규격명):
    # 파일 존재 여부 확인
    if not os.path.exists(file_path):
        print(f"파일을 찾을 수 없습니다: {file_path}")
        return

    open_file(file_path)

    for keyword in keywords:  # 순차적으로 각 키워드 처리
        # 키워드 검색 후 스크린샷 찍기
        screenshot_docx(keyword, 사전규격명)


def screenshot_docx(keyword, 사전규격명):
    # 워드 프로그램 자동화
    try:
        app = pywinauto.Application().connect(path=word_path)  # 워드 프로그램 경로

        word_window = app.window(title_re=".*Word.*")  # 워드 프로그램의 창을 찾기

        # 워드 로딩
        time.sleep(5)

        # 모든 컨트롤 요소들 출력 (child_window)
        word_window.print_control_identifiers()

        # 키워드 검색 (단, 한글 프로그램에서 키워드 검색 기능을 자동화하려면 단축키 활용)
        word_window.type_keys("^f")  # Ctrl+F (검색 단축키)
        logging.info("검색 모달 표시")
        time.sleep(2)

        word_window.type_keys(keyword)  # 검색어 입력
        logging.info("검색어 입력")
        time.sleep(2)

        # 검색
        word_window.type_keys("{ENTER}")
        logging.info("검색 시작")

        # 검색된 텍스트 영역이 활성화되도록 대기
        time.sleep(2)

        capture_count = 0

        while True:
            try:
                # 스크린샷 영역 설정
                x1, y1 = 100, 200  # 좌측 상단 좌표
                width, height = 1800, 800
                x2, y2 = x1 + width, y1 + height  # 우측 하단 좌표 계산

                # 스크린샷 찍기
                screenshot = pyautogui.screenshot(region=(x1, y1, width, height))
                capture_count += 1
                screenshot_file = os.path.join(
                    screenshot_dir, f"{사전규격명}_{keyword}_{capture_count}.png"
                )
                screenshot.save(screenshot_file)
                print(f"검색 결과 {capture_count} 캡처 완료: {screenshot_file}")

                # 다음 검색 결과로 이동
                word_window.set_focus()  # 포커스를 해당 영역으로 이동
                word_window.type_keys("{ENTER}")  # 다음 검색 결과
                time.sleep(2)  # 다음 결과가 로드되도록 대기

            except Exception as e:
                print(f"검색 결과 끝 또는 오류: {e}")
                break

        print(f"총 {capture_count}개의 검색 결과 캡처 완료.")

    except Exception as e:
        print(f"워드 파일 처리 중 오류 발생: {e}")


# ============================================== 엑셀 함수 ==============================================


# 다운로드된 엑셀 파일을 열고, 키워드를 검색하여 스크린샷을 찍는 함수 호출
def handle_xlsx_file(file_path, keywords, 사전규격명):
    # 파일 존재 여부 확인
    if not os.path.exists(file_path):
        print(f"파일을 찾을 수 없습니다: {file_path}")
        return

    open_file(file_path)

    for keyword in keywords:  # 순차적으로 각 키워드 처리
        # 키워드 검색 후 스크린샷 찍기
        screenshot_xlsx(keyword, 사전규격명)


def screenshot_xlsx(keyword, 사전규격명):
    try:
        app = pywinauto.Application().connect(path=excel_path)  # 엑셀 프로그램 경로

        excel_window = app.window(title_re=".*Excel.*")  # 엑셀 프로그램의 창을 찾기

        # 엑셀 로딩
        time.sleep(7)

        # 모든 컨트롤 요소들 출력 (child_window)
        excel_window.print_control_identifiers()

        # 키워드 검색 (단, 한글 프로그램에서 키워드 검색 기능을 자동화하려면 단축키 활용)
        excel_window.type_keys("^f")  # Ctrl+F (검색 단축키)
        logging.info("검색 모달 표시")
        time.sleep(2)

        excel_window.type_keys(keyword)  # 검색어 입력
        logging.info("검색어 입력")
        time.sleep(2)

        # 검색
        excel_window.type_keys("{ENTER}")
        logging.info("검색 시작")

        # 검색된 텍스트 영역이 활성화되도록 대기
        time.sleep(2)

        capture_count = 0

        while True:
            try:
                # 스크린샷 영역 설정
                x1, y1 = 100, 200  # 좌측 상단 좌표
                width, height = 1800, 800
                x2, y2 = x1 + width, y1 + height  # 우측 하단 좌표 계산

                # 스크린샷 찍기
                screenshot = pyautogui.screenshot(region=(x1, y1, width, height))
                capture_count += 1
                screenshot_file = os.path.join(
                    screenshot_dir, f"{사전규격명}_{keyword}_{capture_count}.png"
                )
                screenshot.save(screenshot_file)
                print(f"검색 결과 {capture_count} 캡처 완료: {screenshot_file}")

                # 다음 검색 결과로 이동
                excel_window.set_focus()  # 포커스를 해당 영역으로 이동
                excel_window.type_keys("{ENTER}")  # 다음 검색 결과
                time.sleep(2)  # 다음 결과가 로드되도록 대기

            except Exception as e:
                print(f"검색 결과 끝 또는 오류: {e}")
                break

        print(f"총 {capture_count}개의 검색 결과 캡처 완료.")

    except Exception as e:
        print(f"엑셀 파일 처리 중 오류 발생: {e}")


# ============================================== ZIP 함수 ==============================================


# 다운로드된 zip 폴더 압축 해제
def handle_ZIP(file_path, keywords, 사전규격명):
    # 파일 존재 여부 확인
    if not os.path.exists(file_path):
        print(f"파일을 찾을 수 없습니다: {file_path}")
        return

    # 다운로드 중인 .crdownload 파일 무시
    if file_extension == "crdownload":
        print(f"다운로드가 완료되지 않은 파일입니다: {file_path}")
        return

    extract_folder = os.path.join(download_dir, f"{사전규격명} zip 압축 해제 폴더")
    if not os.path.exists(extract_folder):
        os.makedirs(extract_folder)

    # ZIP 파일 압축 해제
    extract_zip(file_path, extract_folder)

    # 압축 파일 삭제
    delete_zip_file(file_path)

    # 압축 해제된 파일 확장자별 처리
    open_extracted_files(extract_folder, keywords, 사전규격명)


# ZIP 파일을 지정된 폴더로 추출
def extract_zip(file_path, extract_folder):
    try:
        # ZIP 파일 열기
        with zipfile.ZipFile(file_path, "r") as zip_ref:
            # 압축 해제
            zip_ref.extractall(extract_folder)
            print(f"ZIP 파일이 {extract_folder}에 성공적으로 압축 해제되었습니다.")
    except zipfile.BadZipFile:
        print(f"잘못된 ZIP 파일: {file_path}")
    except Exception as e:
        print(f"ZIP 파일 해제 중 오류 발생: {str(e)}")


# 압축 해제된 파일 처리
def open_extracted_files(extract_folder, keywords, 사전규격명):
    # 압축 해제된 폴더 내 파일 리스트 가져오기
    extracted_files = os.listdir(extract_folder)

    if not extracted_files:
        print(f"압축 해제된 폴더에 파일이 없습니다: {extract_folder}")
        return

    for file_name in extracted_files:
        file_path = os.path.join(extract_folder, file_name)

        # 파일 존재 여부 확인
        if not os.path.isfile(file_path):
            print(f"디렉토리 내부에 파일이 아닙니다: {file_name}")
            continue

        # 파일 확장자 추출
        file_extension = file_name.lower().split(".")[-1]

        # 파일 확장자별 처리
        if file_extension == "hwp":
            handle_hwp_file(file_path, keywords, 사전규격명)
        elif file_extension == "hwpx":
            handle_hwpx_file(file_path, keywords, 사전규격명)
        elif file_extension == "pdf":
            handle_pdf_file(file_path, keywords, 사전규격명)
        elif file_extension == "docx":
            handle_docx_file(file_path, keywords, 사전규격명)
        elif file_extension == "xlsx":
            handle_xlsx_file(file_path, keywords, 사전규격명)
        else:
            print(f"지원되지 않는 파일 형식입니다: {file_name}")


# 원본 ZIP 파일 삭제
def delete_zip_file(zip_path):
    try:
        if os.path.exists(zip_path):
            os.remove(zip_path)
            time.sleep(2)
            print(f"원본 ZIP 파일이 성공적으로 삭제되었습니다: {zip_path}")
        else:
            print(f"삭제하려는 ZIP 파일이 존재하지 않습니다: {zip_path}")
    except Exception as e:
        print(f"ZIP 파일 삭제 중 오류 발생: {str(e)}")


# 나라장터 페이지로 이동
driver.get("https://www.g2b.go.kr")
logging.info("나라장터 페이지로 이동 완료")

# 나라장터 페이지 로드 완료될 때까지 sleep 주기
time.sleep(8)

# 창 최대화
driver.maximize_window()

# 팝업 닫기 호출 (조건부 처리)
popups = [
    "#mf_wfm_container_wq_uuid_877_wq_uuid_884_poupR23AB0000013489_close",
    "#mf_wfm_container_wq_uuid_877_wq_uuid_884_poupR23AB0000013478_close",
    "#mf_wfm_container_wq_uuid_877_wq_uuid_884_poupR23AB0000013473_close",
    "#mf_wfm_container_wq_uuid_877_wq_uuid_884_poupR23AB0000013415_close",
    "#mf_wfm_container_wq_uuid_877_wq_uuid_884_poupR23AB0000013414_close",
]

for popup_selector in popups:
    close_popup(popup_selector)
    time.sleep(1)

# 발주 메뉴 클릭
ordering = "mf_wfm_gnb_wfm_gnbMenu_genDepth1_0_btn_menuLvl1_span"
ordering_click = driver.find_element(By.ID, ordering)
ordering_click.click()
logging.info("발주 메뉴 클릭")
time.sleep(1)

# 발주목록 소메뉴 클릭
ordering_list = "#mf_wfm_gnb_wfm_gnbMenu_genDepth1_0_genDepth2_0_btn_menuLvl2"
ordering_list_click = driver.find_element(By.CSS_SELECTOR, ordering_list)
ordering_list_click.click()
logging.info("발주목록 소메뉴 클릭")
time.sleep(5)

# 사전규격공개 옵션 선택
pre_specification = "#mf_wfm_container_radSrchTy_input_1"
pre_specification_click = driver.find_element(By.CSS_SELECTOR, pre_specification)
pre_specification_click.click()
logging.info("검색 유형 사전규격공개 옵션 선택")
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
logging.info("기존 진행일자 시작일 제거 완료")
time.sleep(1)

# 진행일자 시작일 입력
start_date_click.send_keys(yesterday_str)
logging.info(f"시작일 {yesterday_str} 입력 완료")
time.sleep(1)

# 진행일자 종료일 input박스 클릭
end_date_xpath = "//input[@type='text' and contains(@id, 'ibxEndDay')]"
end_date_click = driver.find_element(By.XPATH, end_date_xpath)
end_date_click.click()
time.sleep(1)

# 진행일자 종료일 기존 값 제거
end_date_click.clear()
logging.info("기존 진행일자 종료일 제거 완료")
time.sleep(1)

# 진행일자 종료일 입력
end_date_click.send_keys(yesterday_str)
logging.info(f"종료일 {yesterday_str} 입력 완료")
time.sleep(1)

# 상세 조건 펼치기
detail = "[id$='_btnSearchToggle']"
detail_click = driver.find_element(By.CSS_SELECTOR, detail)
detail_click.click()
logging.info("상세 조건 펼치기 완료")

# 업무구분 일반용역 클릭
work1 = "#mf_wfm_container_chkRqdcBsneSeCd_input_2"
work1_click = driver.find_element(By.CSS_SELECTOR, work1)
work1_click.click()
logging.info("업무구분 일반용역 선택")
time.sleep(1)

# 업무구분 기술영역 클릭
work2 = "#mf_wfm_container_chkRqdcBsneSeCd_input_3"
work2_click = driver.find_element(By.CSS_SELECTOR, work2)
work2_click.click()
logging.info("업무구분 기술영역 선택")
time.sleep(1)

# 사업명 입력 박스 클릭
search_box = "#mf_wfm_container_txtBizNm"
search_box_click = driver.find_element(By.CSS_SELECTOR, search_box)
search_box_click.click()
time.sleep(1)

# 검색 키워드
search_keywords = ["dfdvrg", "시스템", "울산다운", "정보", "리포팅"]

# 파일 내 검색 키워드
file_search_keywords = ["정보", "시스템", "리포트", "Report", "전자문서"]

for search_word in search_keywords:

    # 사업명 입력
    search_box_click.send_keys(search_word)
    logging.info(f"검색 박스에 {search_word} 입력")
    time.sleep(1)

    # 검색 버튼 클릭
    search_button = "#mf_wfm_container_btnS0001"
    search_button_click = driver.find_element(By.CSS_SELECTOR, search_button)
    search_button_click.click()
    logging.info(f"{search_word} 검색 시작")
    time.sleep(1)

    # 리스트에 항목 있는지 확인 (display none을 확인)
    tbody_id = "mf_wfm_container_gridView1_body_tbody"
    time.sleep(1)
    rows = driver.find_elements(By.XPATH, f"//*[@id='{tbody_id}']/tr")
    time.sleep(1)

    # 검색 결과가 없으면 다음 키워드로 계속
    if not any(row.value_of_css_property("display") != "none" for row in rows):
        logging.info(f"'{search_word}' 검색 결과가 없습니다. 다음 키워드로 넘어갑니다.")
        time.sleep(1)
        search_box_click.click()
        time.sleep(1)
        search_box_click.clear()
        time.sleep(2)
        continue  # 검색 결과가 없으면 다음 키워드로 넘어감

    # 현재 작업 중인 항목 인덱스 추적
    current_index = 0

    while current_index < len(rows):
        row = rows[current_index]
        try:
            if "w2grid_hidedRow" in row.get_attribute("class"):
                logging.info(
                    f"'{search_word}'의 검색 결과 리스트 탐색을 완료했습니다. 다음 키워드로 넘어갑니다."
                )
                break  # 해당 키워드로 검색을 종료하고, 다음 키워드로 넘어감
            # 각 row에서 링크를 찾기
            link = row.find_element(By.CSS_SELECTOR, "td a")

            # 링크가 클릭 가능할 때까지 대기
            WebDriverWait(driver, 10).until(EC.element_to_be_clickable(link))

            # 링크 클릭
            link.click()
            logging.info(
                f"리스트에서 {current_index+1}번째 항목의 상세규격정보 페이지로 이동"
            )

            # 새 페이지 로드 대기
            try:
                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located(
                        (By.CSS_SELECTOR, "#mf_wfm_cntsHeader_spnHeaderTitle")
                    )
                )
                logging.info("새 페이지 로드 완료")
            except TimeoutException:
                logging.warning("새 페이지 로드 실패")
                current_index += 1
                continue  # 새 페이지 로드가 실패한 경우 다음 항목으로 넘어감

            time.sleep(2)

            # 첨부파일 여부 확인
            try:
                no_file = driver.find_element(
                    By.XPATH, "//*[contains(@id, '_grdFile_noresult')]"
                )
                time.sleep(1)
                if no_file.is_displayed():
                    time.sleep(2)
                    # 첨부파일은 없지만 사전규격 상세정보 URL이 존재할 때 데이터 추출하기
                    data = extract_data(driver)
                    if data:
                        # 사전규격상세정보_URL이 'N/A'가 아닌 경우에만 엑셀에 저장
                        if data["사전규격상세정보_URL"] != "N/A":
                            logging.info(f"추출된 데이터: {data}")
                            save_to_excel(data)  # 엑셀 파일에 저장
                            time.sleep(1)
                        else:
                            logging.warning("사전규격상세정보 URL이 존재하지 않습니다.")
                            time.sleep(1)
                    else:
                        logging.warning("데이터 추출 실패")
                        time.sleep(1)

                    logging.info("첨부파일이 없습니다. 이전 페이지로 이동합니다.")
                    driver.back()  # 이전 페이지로 이동
                    time.sleep(1)

                    # 페이지가 로드된 후 다시 rows 가져오기
                    WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located(
                            (By.ID, "mf_wfm_container_gridView1_body_tbody")
                        )
                    )
                    rows = driver.find_elements(
                        By.XPATH, f"//*[@id='{tbody_id}']/tr"
                    )
                    current_index += 1
                    time.sleep(1)
                    continue  # 다음 항목으로 넘어감
            except Exception as e:
                logging.info("첨부파일이 있는 것으로 판단됩니다. 계속 진행합니다.")

            # 데이터 추출
            data = extract_data(driver)
            if data:
                logging.info(f"추출된 데이터: {data}")
                time.sleep(1)
            else:
                logging.warning("데이터 추출 실패")
                time.sleep(1)

            # 스크롤 조건: 첨부파일이 화면에 보일 때까지
            if scroll_until_element_visible(driver):
                logging.info("첨부파일 영역으로 이동")
            else:
                logging.warning("첨부파일을 찾지 못했습니다.")

            # 전체 선택 체크박스 클릭
            checkbox = driver.find_element(
                By.CSS_SELECTOR, "[id*='_header__column1_checkboxLabel__id']"
            )
            checkbox.click()
            logging.info("모든 첨부파일을 선택")
            time.sleep(2)

            # 파일 다운로드
            logging.info("파일 다운로드 시작")
            download_button = driver.find_element(
                By.CSS_SELECTOR, "[id*='_btnFileDown']"
            )
            time.sleep(1)
            download_button.click()
            logging.info("파일 다운로드 완료")
            time.sleep(1)

            # 다운로드된 최신 파일 찾기
            logging.info("다운로드된 파일 찾는 중")
            latest_file = get_latest_downloaded_file(download_dir)
            time.sleep(2)

            # 파일 열기
            if latest_file and data:
                # 추출한 data에서 사전규격명 가져오기 (파일명에 사용)
                사전규격명 = data.get("사전규격명", "N/A")

                logging.info(f"최근 다운로드된 파일: {latest_file}")
                # handle_file(latest_file)
                file_extension = latest_file.lower().split(".")[-1]  # 확장자 확인

                if file_extension == "hwp":  # HWP 파일인 경우
                    logging.info("한글 파일 처리 시작")
                    handle_hwp_file(latest_file, file_search_keywords, 사전규격명)
                elif file_extension == "hwpx":  # HWPX 파일
                    logging.info("HWPX 파일 처리 시작")
                    handle_hwpx_file(latest_file, file_search_keywords, 사전규격명)
                elif file_extension == "pdf":  # PDF 파일
                    logging.info("PDF 파일 처리 시작")
                    handle_pdf_file(latest_file, file_search_keywords, 사전규격명)
                elif file_extension == "docx":  # Word 파일
                    logging.info("Word 파일 (DOCX) 처리 시작")
                    handle_docx_file(latest_file, file_search_keywords, 사전규격명)
                elif file_extension == "xlsx":  # Excel 파일
                    logging.info("Excel 파일 (XLSX) 처리 시작")
                    handle_xlsx_file(latest_file, file_search_keywords, 사전규격명)
                elif file_extension == "zip":  # ZIP 폴더
                    logging.info("ZIP 폴더 처리 시작")
                    handle_ZIP(latest_file, file_search_keywords, 사전규격명)
            else:
                logging.warning("다운로드된 파일이 없습니다.")

            # 파일 처리 후 이전 페이지로 돌아가기
            driver.back()
            time.sleep(1)

            # 페이지가 로드된 후 다시 rows 가져오기
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "mf_wfm_container_gridView1_body_tbody"))
            )
            rows = driver.find_elements(By.XPATH, f"//*[@id='{tbody_id}']/tr")
            time.sleep(1)

            # 다음 항목 처리를 위해 current_index 증가
            current_index += 1

            # 다음 row가 있는지 확인 후 처리
            if current_index < len(rows):
                time.sleep(1)
                continue  # 다음 항목 처리
            else:
                try:
                    # 현재 선택된 페이지 확인
                    current_page = driver.find_element(By.CLASS_NAME, "w2pageList_label_selected")
                    current_page_number = int(current_page.text)

                    # 다음 페이지 버튼 찾기
                    next_page_number = current_page_number + 1
                    try:
                        next_page_button = driver.find_element(By.ID, f"mf_wfm_container_pagelist_page_{next_page_number}")
                    except:
                        logging.info("다음 페이지 없음. 모든 검색 완료.")
                        break  # 더 이상 페이지가 없으면 종료

                    # 다음 페이지로 이동
                    next_page_button.click()
                    logging.info(f"{next_page_number} 페이지로 이동 중...")
                    time.sleep(2)

                    WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.ID, tbody_id))
                    )

                    rows = driver.find_elements(By.XPATH, f"//*[@id='{tbody_id}']/tr")
                    current_index = 0

                except Exception as e:
                    logging.warning(f"페이지 이동 실패: {e}")
                    break

        except StaleElementReferenceException:
            logging.warning("Stale element encountered. 현재 row를 건너뜁니다.")
            current_index += 1  # Stale element가 발생하면 건너뛰고 계속 진행
        except Exception as e:
            logging.error(f"예상치 못한 오류 발생: {e}")
            current_index += (
                1  # 오류가 발생하면 해당 항목을 건너뛰고 다음 항목으로 넘어감
            )
    # 검색어 입력 박스를 다시 찾기 (페이지가 갱신될 수 있음)
    search_box_click = driver.find_element(By.CSS_SELECTOR, search_box)
    time.sleep(1)
    search_box_click.clear()
    time.sleep(1)


# 모든 검색 키워드를 처리한 후 종료
logging.info("모든 검색이 완료되었습니다. 프로그램을 종료합니다.")
sys.exit()
