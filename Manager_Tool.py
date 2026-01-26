import os
import subprocess
import requests
import sys

# ==============================================================================
# [설정]
# ==============================================================================
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
        
        # 최신 버전 순 정렬
        sorted_versions = sorted(versions, key=lambda x: [int(p) for p in x['version'].split('.')], reverse=True)
        
        # 메이저 버전별 최신 드라이버 URL 추출 (win32 우선, 없으면 win64)
        major_map = {}
        for v in sorted_versions:
            major = v['version'].split('.')[0]
            if major not in major_map:
                driver_url = None
                for d in v['downloads'].get('chromedriver', []):
                    if d['platform'] == 'win32':
                        driver_url = d['url']
                        break
                if not driver_url:
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

        # 다운로드 및 저장
        for major_ver, url in major_map.items():
            file_name = f"chromedriver_{major_ver}.exe"
            file_path = os.path.join(DRIVERS_DIR, file_name)
            
            if os.path.exists(file_path):
                print(f"  Existing: {file_name} (건너뜀)")
                continue
                
            print(f"  ⬇️ Downloading: {file_name}...")
            
            # (간이 다운로드 로직: 실제로는 zip 해제가 필요할 수 있으나, 
            #  기존에 파일 생성 로직이 작동한다고 가정하고 requests로 바로 저장하는 예시입니다)
            try:
                # 주의: 실제 URL은 zip 파일일 확률이 높으므로, 
                # zip을 받아서 압축을 풀고 exe만 꺼내는 로직이 정석입니다.
                # 여기서는 '목록 갱신' 기능 자체보다는 'Git 업로드' 기능 수정에 집중합니다.
                pass 
                
            except Exception as e:
                print(f"    다운로드 실패: {e}")

    except Exception as e:
        print(f"❌ 드라이버 목록 갱신 실패: {e}")

def sync_to_github():
    print("\n🚀 GitHub 동기화 시작 (Drivers 폴더만)...")
    
    # 1. 변경사항 추가 (Add) - ★수정된 부분★
    # 온점(.) 대신 폴더명(drivers)을 적어서 이 폴더만 올립니다.
    print(f"📦 '{DRIVERS_DIR}' 폴더만 담는 중 (git add)...")
    subprocess.call(f"git add {DRIVERS_DIR}", shell=True)
    
    # 2. 커밋 (Commit)
    print("📝 기록 남기는 중 (git commit)...")
    subprocess.call('git commit -m "Update ChromeDriver list only"', shell=True)
    
    # 3. 원격 변경사항 가져오기 (Pull) - 충돌 방지
    print("🔄 서버 상태 확인 중 (git pull)...")
    subprocess.call("git pull origin main --no-edit", shell=True)

    # 4. 업로드 (Push)
    print("📤 깃허브로 업로드 중 (git push)...")
    push_result = subprocess.call("git push origin main", shell=True)
    
    if push_result == 0:
        print("\n✅ 드라이버 업데이트 완료! (다른 파일은 건드리지 않았습니다)")
    else:
        print("\n❌ 업로드 실패.")

if __name__ == "__main__":
    # 1. 드라이버 다운로드 (필요시 주석 해제)
    # get_latest_drivers()
    
    # 2. 깃허브 동기화 (드라이버 폴더만)
    sync_to_github()