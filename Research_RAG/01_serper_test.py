import os
import requests
import json
from dotenv import load_dotenv

# 환경 변수 로드
load_dotenv()
SERPER_API_KEY = os.getenv("SERPER_API_KEY")

def search_riss_paper(query):
    url = "https://google.serper.dev/search"

    # RISS 사이트 내에서만 검색하도록 검색어 조작
    payload = json.dumps({
        "q": f"site:riss.kr {query}",
        "num": 5,
        "gl": "kr",  # 한국 지역 설정
        "hl": "ko"   # 한국어 설정
    })

    headers = {
        'X-API-KEY': SERPER_API_KEY,
        'Content-Type': 'application/json'
    }

    print(f"🔍 Serper로 검색 중: '{query}' (RISS 한정)...")

    try:
        response = requests.request("POST", url, headers=headers, data=payload)

        if response.status_code == 200:
            results = response.json()
            organic = results.get("organic", [])

            print(f"\n✅ 검색 성공! (총 {len(organic)}건)\n")

            for i, item in enumerate(organic, 1):
                print(f"[{i}] {item.get('title')}")
                print(f"    🔗 {item.get('link')}")
                print(f"    📝 {item.get('snippet')[:50]}...")
                print("-" * 40)
        else:
            print(f"❌ 오류: {response.text}")

    except Exception as e:
        print(f"❌ 실행 오류: {e}")

if __name__ == "__main__":
    search_riss_paper("미군정 영화 검열")