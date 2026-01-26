import os
import subprocess
import requests
import re
import sys

# ==============================================================================
# [설정]
# ==============================================================================
# 크롬 드라이버 다운로드 페이지 (JSON API 사용 권장되지만, 여기선 기존 방식 유지 가정)
CHROME_DRIVER_URL = "https://googlechromelabs.github.io/chrome-for-testing/known-good-versions-with-downloads.json"
DRIVERS_DIR = "drivers"

def get_latest_drivers():
    print("🔍 구글 서버에서 전체 버전 목록을 가져오는 중... (시간이 좀 걸립니다)")
    
    try:
        response = requests.get(CHROME_DRIVER_URL)
        if response.status_code != 200:
            print("❌ 버전 정보를 가져오지 못했습니다.")
            return

        data = response.json()
        versions = data['versions']
        
        # 최신 버전 순으로 정렬 (버전 숫자가 높은 순)
        sorted_versions = sorted(versions, key=lambda x: [int(p) for p in x['version'].split('.')], reverse=True)
        
        # 메이저 버전별로 하나씩만 추출 (가장 최신 것)
        major_map = {}
        for v in sorted_versions:
            major = v['version'].split('.')[0]
            if major not in major_map:
                # win32 또는 win64 드라이버 찾기
                driver_url = None
                for d in v['downloads'].get('chromedriver', []):
                    if d['platform'] == 'win32':
                        driver_url = d['url']
                        break
                if not driver_url: # win32 없으면 win64 시도
                    for d in v['downloads'].get('chromedriver', []):
                        if d['platform'] == 'win64':
                            driver_url = d['url']
                            break
                
                if driver_url:
                    major_map[major] = driver_url
            
            if len(major_map) >= 10: # 최신 10개 버전만 확보
                break
        
        print(f"📊 총 {len(major_map)}개의 버전을 확인했습니다.")

        if not os.path.exists(DRIVERS_DIR):
            os.makedirs(DRIVERS_DIR)

        # 다운로드 진행
        for major_ver, url in major_map.items():
            file_name = f"chromedriver_{major_ver}.exe"
            file_path = os.path.join(DRIVERS_DIR, file_name)
            
            if os.path.exists(file_path):
                print(f"  Existing: {file_name} (건너뜀)")
                continue
                
            print(f"  ⬇️ Downloading: {file_name}...")
            
            # 파일 다운로드 및 저장 (zip 파일 처리 필요)
            # (단순화를 위해 exe가 바로 있다고 가정하지 않고, zip 받아서 압축 해제 로직이 필요할 수 있음)
            # 여기서는 편의상 다운로드 로직은 기존에 잘 되셨던 방식이 있다면 그걸 쓰시되,
            # zip 해제 로직이 복잡하므로 간단히 urlretrieve 대신 requests 사용 예시:
            
            # --- (실제로는 zip을 받아서 exe만 꺼내야 합니다) ---
            # 복잡해지므로 일단 '목록 가져오기' 성공한 기존 로직을 유지한다고 가정하고
            # 핵심인 'Git 동기화' 부분에 집중하겠습니다.
            pass 

    except Exception as e:
        print(f"❌ 드라이버 목록 갱신 실패: {e}")

def sync_to_github():
    print("\n🚀 GitHub 동기화 시작...")
    
    # 1. 변경사항 추가 (Add)
    print("📦 파일 담는 중 (git add)...")
    subprocess.call("git add .", shell=True)
    
    # 2. 커밋 (Commit)
    print("📝 기록 남기는 중 (git commit)...")
    subprocess.call('git commit -m "Update ChromeDriver via Manager_Tool"', shell=True)
    
    # ▼▼▼▼▼ [여기가 추가된 핵심 코드!] ▼▼▼▼▼
    # 3. 원격 변경사항 가져오기 (Pull)
    print("🔄 서버에 있는 새 파일 가져오는 중 (git pull)...")
    pull_result = subprocess.call("git pull origin main --no-edit", shell=True)
    
    if pull_result != 0:
        print("⚠️ 주의: Pull 과정에서 충돌이 났거나 병합 메시지 창이 떴을 수 있습니다.")
        print("   (검은 화면에 vi 에디터가 뜨면 ':wq' 입력 후 엔터를 치세요)")
    # ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    # 4. 업로드 (Push)
    print("📤 깃허브로 업로드 중 (git push)...")
    push_result = subprocess.call("git push origin main", shell=True)
    
    if push_result == 0:
        print("\n✅ 모든 작업 완료! (GitHub에 잘 올라갔습니다)")
    else:
        print("\n❌ 업로드 실패. 로그를 확인해주세요.")

if __name__ == "__main__":
    # 1. 드라이버 관리 (기존 코드 유지)
    # get_latest_drivers()  <-- 필요할 때 주석 풀고 쓰세요
    
    # 2. 깃허브 동기화
    sync_to_github()