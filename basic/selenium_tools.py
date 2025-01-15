import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

import chromedriver_autoinstaller

chromedriver_autoinstaller.install()

driver = webdriver.Chrome()

# 1. Navigation 관련 툴
# get, back, forward, refresh

# # 1-1. get() 원하는 페이지로 이동하는 함수
# driver.get("https://www.naver.com")
# time.sleep(1)
# driver.get("https://www.google.com")

# # 1-2. back() - 뒤로 가기
# driver.back()
# time.sleep(2)

# # 1-3. forward() - 앞으로 가기
# driver.forward()
# time.sleep(2)

# # 1.4. refresh() - 페이지 새로고침
# driver.refresh()
# time.sleep(2)


# 2. browser infomation
driver.get("https://www.naver.com")
time.sleep(2)
# 2-1. title - 웹 사이트의 타이틀 가지고 옴
title = driver.title
print(title, "이 타이틀이다")

# 2-2. current_url - 주소창을 그대로 가지고 옴옴
current = driver.current_url
print(current, "가 현재 주소임")

# 3. Driver Wait
# 3-1. 3초 때 로딩이 끝나서, element가 찾아짐
# 3-2. 30초까지는 기다림
# 3-3. 30초 넘어가면 에러 던짐.
selector = "#topPayArea"
try:
    WebDriverWait(driver, 5).until(EC.presence_of_element_located(
    By.CSS_SELECTOR, selector
))
except:
    print("예외 발생, 예외 처리 코드 실행행")
print("로딩 완료")
print("다음 코드 실행")

input()