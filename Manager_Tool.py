import os
import requests
import zipfile
import io
import subprocess
import shutil
from datetime import datetime

# ==============================================================================
# [설정] 
# ==============================================================================
# 1. 크롬 드라이버 저장할 폴더명 (없으면 현재 위치에 저장하려면 "" 로 두세요)
TARGET_DIR = "drivers" 

# 2. 깃허브 커밋 메시지
COMMIT_MSG = "Update ChromeDriver automatically"

def get_latest_stable_version():
    """구글 API를 통해 최신 Stable 버전을 가져옵니다."""
    url = "https://googlechromelabs.github.io/chrome-for-testing/last-known-good-versions-with-downloads.json"
    try:
        res = requests.get(url)
        data = res.json()
        version = data['channels']['Stable']['version']
        
        # 다운로드 URL 찾기 (win64 기준)
        downloads = data['channels']['Stable']['downloads']['chromedriver']
        download_url = ""
        for item in downloads:
            if item['platform'] == 'win64':
                download_url = item['url']
                break
        
        return version.split('.')[0], download_url # (예: "121", "https://...")
    except Exception as e:
        print(f"❌ 버전 확인 실패: {e}")
        return None, None

def download_and_process_driver():
    # 1. 최신 버전 정보 가져오기
    print("🔍 최신 크롬 드라이버 정보를 확인 중...")
    major_ver, url = get_latest_stable_version()
    
    if not major_ver:
        return

    target_filename = f"chromedriver_{major_ver}.exe"
    
    # 폴더 처리
    if TARGET_DIR and not os.path.exists(TARGET_DIR):
        os.makedirs(TARGET_DIR)
    
    final_path = os.path.join(TARGET_DIR, target_filename) if TARGET_DIR else target_filename

    # 이미 파일이 있는지 확인
    if os.path.exists(final_path):
        print(f"✅ 이미 최신 버전({major_ver}) 파일이 존재합니다: {target_filename}")
        # 파일이 있어도 깃허브에 안 올라가 있을 수 있으니 업로드 단계로 넘어감
    else:
        # 2. 다운로드 및 압축 해제
        print(f"⬇️ 다운로드 시작 (v{major_ver})...")
        try:
            res = requests.get(url)
            with zipfile.ZipFile(io.BytesIO(res.content)) as z:
                # 압축 파일 내부 구조: chromedriver-win64/chromedriver.exe
                # 파일만 쏙 빼서 저장
                for file_info in z.infolist():
                    if file_info.filename.endswith("chromedriver.exe"):
                        with z.open(file_info) as source, open(final_path, "wb") as target:
                            shutil.copyfileobj(source, target)
                        break
            print(f"✅ 파일 생성 완료: {final_path}")
        except Exception as e:
            print(f"❌ 다운로드/저장 실패: {e}")
            return

    # 3. GitHub에 푸시 (Git 명령어 사용)
    push_to_github(final_path)

def push_to_github(filepath):
    print("🚀 GitHub 업로드 시작...")
    
    # 현재 폴더가 git 리포지토리인지 확인
    if not os.path.exists(".git"):
        print("❌ 현재 폴더에 .git이 없습니다. Git 리포지토리 루트에서 실행해주세요.")
        return

    try:
        # git add
        subprocess.run(f'git add "{filepath}"', shell=True, check=True)
        print(f"   - Staged: {filepath}")
        
        # git commit (변경사항이 없으면 에러 날 수 있으므로 try 처리)
        try:
            subprocess.run(f'git commit -m "{COMMIT_MSG}"', shell=True, check=True)
            print("   - Committed.")
        except:
            print("   - (변경사항 없음 또는 이미 커밋됨)")

        # git push
        subprocess.run("git push", shell=True, check=True)
        print("🎉 GitHub Push 완료! 성공적으로 업로드되었습니다.")
        
    except subprocess.CalledProcessError as e:
        print(f"❌ Git 명령 실패: {e}")

if __name__ == "__main__":
    # 로그 파일에 실행 기록 남기기 (log.txt)
    with open("manager_log.txt", "a", encoding="utf-8") as f:
        f.write(f"[{datetime.now()}] 실행 시작\n")
    
    try:
        download_and_process_driver()
        msg = "성공"
    except Exception as e:
        msg = f"실패: {e}"
        
    with open("manager_log.txt", "a", encoding="utf-8") as f:
        f.write(f"[{datetime.now()}] 실행 종료 - {msg}\n\n")
    
    # input() 제거! (자동 종료되도록)