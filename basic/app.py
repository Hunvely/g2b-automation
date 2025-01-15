import time
from selenium import webdriver
from selenium.webdriver.common.by import By
import chromedriver_autoinstaller

chromedriver_autoinstaller.install()

driver = webdriver.Chrome()
# 1. 웹 브라우저 주소창을 컨트롤하기 - driver.get
driver.get("http://www.naver.com")

time.sleep(3)

# 2-1. 요소를 찾아서 Copy 해옴. 실제 웹 브라우저 + 개발자 도구
css_selector = "#shortcutArea"

# 2-2. 찾아온 요소를 find_element로 가져와 변수에 선언
navi = driver.find_element(By.CSS_SELECTOR, css_selector)

# 3-1. 데이터 가져오기
print(navi.text)

# 3-2. 요소를 클릭하기 (action)
navi.click()

input()