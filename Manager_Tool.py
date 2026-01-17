import os
import requests
import zipfile
import io
import subprocess
import shutil
from datetime import datetime
import urllib3

# ==============================================================================
# [설정] 
# ==============================================================================
TARGET_DIR = "drivers"
COMMIT_MSG = "Update ChromeDriver (Latest 10 versions)"
MAX_VERSIONS = 10  # 최근 10개 버전까지 수집

# 보안 경고 무시
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

def get_driver_links():
    """
    구글의 전체 버전 목록을 뒤져서, 
    각 메이저 버전(144, 143, 142...)별로 '가장 최신 빌드' 하나씩을 뽑아냅니다.
    """
    print("🔍 구글 서버에서 전체 버전 목록을 가져오는 중... (시간이 좀 걸립니다)")
    
    # 전체 버전 정보가 있는 JSON (용량이 좀 큽니다)
    url = "https://googlechromelabs.github.io/chrome-for-testing/known-good-versions-with-downloads.json"
    
    try:
        res = requests.get(url, verify=False)
        data = res.json()
        versions = data['versions']
        
        # 최신 버전순으로 정렬 (버전 숫자가 높은 게 위로 오게)
        # 버전 문자열(144.0.1234.5)을 숫자 리스트로 변환해서 정렬
        versions.sort(key=lambda x: [int(p) for p in x['version'].split('.')], reverse=True)
        
        collected_drivers = {} # { 144: "다운로드주소", 143: "다운로드주소" ... }
        
        for v in versions:
            version_str = v['version']
            major_ver = int(version_str.split('.')[0])
            
            # 이미 수집한 메이저 버전이면 패스 (우리는 각 버전의 '최신'만 필요하므로)
            if major_ver in collected_drivers:
                continue
                
            # win64 드라이버가 있는지 확인
            if 'chromedriver' in v['downloads']:
                for item in v['downloads']['chromedriver']:
                    if item['platform'] == 'win64':
                        collected_drivers[major_ver] = item['url']
                        break
            
            # 목표 개수(10개) 채웠으면 중단
            if len(collected_drivers) >= MAX_VERSIONS:
                break
                
        return collected_drivers

    except Exception as e:
        print(f"❌ 버전 목록 확보 실패: {e}")
        return {}

def download_and_save():
    # 1. 다운로드 할 목록 가져오기
    drivers_map = get_driver_links()
    
    if not drivers_map:
        print("❌ 다운로드 할 드라이버가 없습니다.")
        return

    # 폴더 생성
    if not os.path.exists(TARGET_DIR):
        os.makedirs(TARGET_DIR)

    print(f"📊 총 {len(drivers_map)}개의 버전을 다운로드합니다. (최신 {max(drivers_map.keys())} ~ 과거 {min(drivers_map.keys())})")

    # 2. 하나씩 다운로드
    for major_ver, url in drivers_map.items():
        filename = f"chromedriver_{major_ver}.exe"
        save_path = os.path.join(TARGET_DIR, filename)
        
        # 파일이 이미 있으면 건너뛰기 (불필요한 트래픽 방지)
        if os.path.exists(save_path):
            print(f"  Existing: {filename} (건너뜀)")
            continue
            
        print(f"  ⬇️ Downloading: {filename} ...")
        
        try:
            res = requests.get(url, verify=False)
            with zipfile.ZipFile(io.BytesIO(res.content)) as z:
                for file_info in z.infolist():
                    if file_info.filename.endswith("chromedriver.exe"):
                        with z.open(file_info) as source, open(save_path, "wb") as target:
                            shutil.copyfileobj(source, target)
                        break
        except Exception as e:
            print(f"  ❌ 실패 ({filename}): {e}")

    # 3. 깃허브 업로드
    push_to_github()

def push_to_github():
    print("\n🚀 GitHub 동기화 시작...")
    
    if not os.path.exists(".git"):
        print("❌ .git 폴더가 없습니다.")
        return

    try:
        # 변경된 모든 파일 담기 (새로 받은 드라이버들)
        subprocess.run('git add .', shell=True, check=True)
        
        # 커밋
        try:
            subprocess.run(f'git commit -m "{COMMIT_MSG}"', shell=True, check=True)
            print("   - 커밋 완료.")
        except:
            print("   - (변경사항 없음)")
            # 변경사항 없어도 push는 시도 (혹시 누락된 게 있을 수 있으니)

        # 업로드
        subprocess.run("git push", shell=True, check=True)
        print("🎉 GitHub Push 완료! 모든 버전이 업로드되었습니다.")
        
    except subprocess.CalledProcessError as e:
        print(f"❌ Git 명령 실패: {e}")

if __name__ == "__main__":
    with open("manager_log.txt", "a", encoding="utf-8") as f:
        f.write(f"[{datetime.now()}] 다중 버전 다운로드 실행\n")
    
    try:
        download_and_save()
    except Exception as e:
        print(f"치명적 오류: {e}")
        with open("manager_log.txt", "a", encoding="utf-8") as f:
            f.write(f"오류 발생: {e}\n")