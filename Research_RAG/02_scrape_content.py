import time
import os
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

def scrape_and_save_riss():
    # 타겟 논문 URL
    url = "https://www.riss.kr/search/detail/DetailView.do?p_mat_type=1a0202e37d52c72d&control_no=76bd82d362ea53fd6aae8a972f9116fb&keyword=%EB%AF%B8%EA%B5%B0%EC%A0%95%EA%B8%B0%20%EC%98%81%ED%99%94%20%EA%B2%80%EC%97%B4"

    print(f"📡 [Selenium] 논문 데이터 수집 중...\n🔗 {url}")

    chrome_options = Options()
    chrome_options.add_argument("--headless") 
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")

    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)

    try:
        driver.get(url)
        wait = WebDriverWait(driver, 10)
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.thesisInfo")))
        
        # 1. [확장] 더보기 버튼 무차별 클릭
        try:
            time.sleep(1.5)
            buttons = driver.find_elements(By.CSS_SELECTOR, "a.moreView, a.btn_more")
            print(f"👉 내용 확장을 위해 버튼 {len(buttons)}개를 클릭합니다.")
            for btn in buttons:
                if btn.is_displayed():
                    driver.execute_script("arguments[0].click();", btn)
                    time.sleep(0.2)
            time.sleep(1) 
        except:
            pass

        # 2. [수집] 제목 및 본문
        # 제목
        h3_tags = driver.find_elements(By.TAG_NAME, "h3")
        title_text = max([t.text for t in h3_tags], key=len).strip() if h3_tags else "제목_없음"

        # 본문 (진공청소기 전략)
        content_text = ""
        try:
            content_text += driver.find_element(By.CSS_SELECTOR, "div.additionalInfo").text + "\n\n"
        except: pass
        
        try:
            content_text += driver.find_element(By.CSS_SELECTOR, "div.text").text
        except: pass
        
        # 본문이 비었을 경우 백업 (body 전체 검색)
        if len(content_text) < 50:
            body_text = driver.find_element(By.TAG_NAME, "body").text
            if "국문초록" in body_text:
                content_text = body_text.split("국문초록")[1]

        # 3. [저장] 파일로 내보내기
        # 파일명에 특수문자가 있으면 에러나므로 정리
        safe_title = "".join([c for c in title_text if c.isalnum() or c in (' ', '_')]).strip()[:30]
        filename = f"paper_{safe_title}.txt"
        
        # Research_RAG 폴더 안에 저장
        save_path = os.path.join("Research_RAG", filename)

        with open(save_path, "w", encoding="utf-8") as f:
            f.write(f"TITLE: {title_text}\n")
            f.write(f"URL: {url}\n")
            f.write("="*40 + "\n\n")
            f.write(content_text)

        print("\n" + "="*60)
        print(f"🎉 [수집 완료] 전체 내용을 파일로 저장했습니다!")
        print(f"📂 저장 위치: {save_path}")
        print("="*60)

    except Exception as e:
        print(f"❌ 오류: {e}")
    finally:
        driver.quit()

if __name__ == "__main__":
    scrape_and_save_riss()