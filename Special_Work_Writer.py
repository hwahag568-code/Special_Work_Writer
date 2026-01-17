import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import threading
import time
from datetime import datetime
import calendar
import os
from selenium.webdriver.common.keys import Keys 
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import subprocess
import sys
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
import socket # [추가] 크롬 실행 여부 확인용
import requests # [추가] 업데이트 체크용 (pip install requests 필요)
import winreg # [추가] 윈도우 레지스트리 접근용
import zipfile # (혹시 zip으로 받을 경우 대비, 지금은 exe라 안씀)

# ==============================================================================
# [설정 1] 자동 업데이트 정보 (여기를 본인 GitHub 주소로 수정하세요!)
# ==============================================================================
CURRENT_VERSION = "1.5"
REPO_URL_VERSION = "https://raw.githubusercontent.com/hwahag568-code/Special_Work_Writer/main/version.txt"
REPO_URL_INSTALLER = "https://raw.githubusercontent.com/hwahag568-code/Special_Work_Writer/main/Update_Work_Writer.exe"

# ==============================================================================
# [설정] 드라이버 저장소 주소 (본인 주소로 변경!)
# ==============================================================================
# 예: https://raw.githubusercontent.com/사용자ID/리포지토리/main/chromedriver_
DRIVER_BASE_URL = "https://raw.githubusercontent.com/hwahag568-code/Special_Work_Writer/main/chromedriver_"

# ==============================================================================
# [설정 2] 사용자 제공 XPath
# ==============================================================================
XPATHS = {
    "CALENDAR_ICON": '//*[@id="ipcSTRT_YMDX_img"]',
    "REASON_INPUT":  '//*[@id="iptAPPL_RMRK"]',
    "SUBMIT_BTN":    '//*[@id="btnSetEMPL_NUMB_center"]',
    "START_H_BTN":   '//*[@id="cmbSTRT_HHXX_button"]',
    "START_M_IPT":   '//*[@id="iptSTRT_MMXX"]',
    "END_H_BTN":     '//*[@id="cmbENDX_HHXX_button"]',
    "END_M_IPT":     '//*[@id="iptENDX_MMXX"]',
    "POPUP_CONFIRM": '//*[@id="btn_confirm"]', 
    "DUPLICATE_MSG": '//*[@id="grpMessage"]'   
}
# ==============================================================================

class SimpleCalendar(tk.Frame):
    def __init__(self, parent, select_callback):
        super().__init__(parent)
        self.select_callback = select_callback
        self.buttons = {}
        self.selected_dates = set()
        self.year = datetime.now().year
        self.month = datetime.now().month
        self.create_widgets()

    def create_widgets(self):
        # 상단 (년/월 이동)
        header = tk.Frame(self)
        header.pack(fill="x", pady=2)
        tk.Button(header, text="<", command=self.prev_month).pack(side="left")
        self.lbl_header = tk.Label(header, text=f"{self.year}년 {self.month}월", font=("Arial", 11, "bold"))
        self.lbl_header.pack(side="left", expand=True)
        tk.Button(header, text=">", command=self.next_month).pack(side="right")

        # [수정] 요일과 날짜를 같은 프레임(grid_frame)에 넣어서 줄 맞춤
        self.grid_frame = tk.Frame(self)
        self.grid_frame.pack()
        self.draw_days()



    def draw_days(self):
            # 기존 내용물 삭제
            for widget in self.grid_frame.winfo_children():
                widget.destroy()
            self.buttons = {}

            # 1. 요일 헤더 (일, 월, 화...)
            days = ["일", "월", "화", "수", "목", "금", "토"]
            for i, d in enumerate(days):
                lbl = tk.Label(self.grid_frame, text=d, width=4, fg="red" if d=="일" else "black")
                lbl.grid(row=0, column=i, padx=1, pady=(0, 5))

            # 2. 날짜 버튼 그리기 (수정됨: 일요일 시작 기준)
            # firstweekday=6 은 일요일을 뜻합니다. (기본값은 0=월요일)
            cal = calendar.Calendar(firstweekday=6).monthdayscalendar(self.year, self.month)
            
            for r, week in enumerate(cal):
                for c, day in enumerate(week):
                    if day == 0:
                        tk.Label(self.grid_frame, text="", width=4).grid(row=r+1, column=c)
                    else:
                        date_str = f"{self.year}-{self.month:02d}-{day:02d}"
                        btn = tk.Button(self.grid_frame, text=str(day), width=4, bg="#f0f0f0",
                                        command=lambda d=date_str: self.toggle_date(d))
                        btn.grid(row=r+1, column=c, padx=1, pady=1)
                        self.buttons[date_str] = btn
                        
                        if date_str in self.selected_dates:
                            btn.config(bg="#3b8ed0", fg="white")

    def toggle_date(self, date_str):
        if date_str in self.selected_dates:
            self.selected_dates.remove(date_str)
            self.buttons[date_str].config(bg="#f0f0f0", fg="black")
        else:
            self.selected_dates.add(date_str)
            self.buttons[date_str].config(bg="#3b8ed0", fg="white")
        self.select_callback(self.selected_dates)

    def prev_month(self):
        self.month -= 1
        if self.month == 0: self.month, self.year = 12, self.year - 1
        self.lbl_header.config(text=f"{self.year}년 {self.month}월")
        self.draw_days()

    def next_month(self):
        self.month += 1
        if self.month == 13: self.month, self.year = 1, self.year + 1
        self.lbl_header.config(text=f"{self.year}년 {self.month}월")
        self.draw_days()

class ChromeDriverUpdater:
    def get_chrome_version(self):
        """내 PC에 설치된 크롬의 메이저 버전(예: 121)을 가져옵니다."""
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Google\Chrome\BLBeacon")
            version, _ = winreg.QueryValueEx(key, "version")
            return version.split('.')[0] # "121.0.6167.85" -> "121"
        except:
            try:
                # 64비트 레지스트리 확인
                key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Google\Chrome\BLBeacon")
                version, _ = winreg.QueryValueEx(key, "version")
                return version.split('.')[0]
            except:
                return None

    def get_driver_version(self, driver_path):
        """현재 가지고 있는 드라이버의 버전을 확인합니다."""
        if not os.path.exists(driver_path):
            return None
        try:
            # 드라이버 실행해서 버전 확인 (--version)
            result = subprocess.check_output([driver_path, "--version"], stderr=subprocess.STDOUT)
            # 결과 예: "ChromeDriver 121.0.6167.85 (..."
            version_str = result.decode('utf-8').split(' ')[1]
            return version_str.split('.')[0] # "121"
        except:
            return None

    def update_driver_if_needed(self, driver_name="chromedriver.exe"):
        """크롬 버전과 드라이버 버전을 비교해서 다르면 다운로드합니다."""
        
        # 1. 크롬 버전 확인
        chrome_ver = self.get_chrome_version()
        if not chrome_ver:
            print("❌ 크롬이 설치되어 있지 않거나 버전을 찾을 수 없습니다.")
            return False # 진행 불가

        # 2. 드라이버 버전 확인
        driver_ver = self.get_driver_version(driver_name)
        
        print(f"🔎 버전 체크 - Chrome: {chrome_ver} / Driver: {driver_ver}")

        # 3. 버전 일치하면 패스
        if chrome_ver == driver_ver:
            print("✅ 드라이버 버전이 일치합니다.")
            return True

        # 4. 불일치 또는 드라이버 없음 -> 다운로드 시도
        print(f"🔄 드라이버 업데이트 필요! (v{chrome_ver} 다운로드 시도)")
        
        # 다운로드 URL 생성 (예: .../chromedriver_121.exe)
        download_url = f"{DRIVER_BASE_URL}{chrome_ver}.exe"
        
        try:
            requests.packages.urllib3.disable_warnings()
            response = requests.get(download_url, verify=False, timeout=10)
            
            if response.status_code == 200:
                # 기존 파일 있으면 삭제 (PermissionError 방지 위해 taskkill 선행 필요할 수 있음)
                if os.path.exists(driver_name):
                    try:
                        os.remove(driver_name)
                    except:
                        subprocess.call(f"taskkill /f /im {driver_name}", shell=True)
                        time.sleep(1)
                        os.remove(driver_name)

                # 새 파일 저장
                with open(driver_name, "wb") as f:
                    f.write(response.content)
                
                print(f"🎉 드라이버 업데이트 완료 (v{chrome_ver})")
                return True
            else:
                print(f"❌ GitHub에 해당 버전(v{chrome_ver}) 드라이버가 없습니다.")
                messagebox.showerror("오류", f"현재 크롬 버전({chrome_ver})에 맞는 드라이버가 서버에 없습니다.\n개발자에게 요청하세요.")
                return False
                
        except Exception as e:
            print(f"❌ 드라이버 다운로드 실패: {e}")
            messagebox.showerror("오류", f"드라이버 업데이트 실패: {e}")
            return False

class AutoWorkApp:
    def __init__(self, root):
        self.root = root
        self.root.title("특근 입력기")
        
        window_width = 430
        window_height = 650
        screen_width = root.winfo_screenwidth()
        x_pos = screen_width - window_width - 200
        y_pos = 100
        self.root.geometry(f"{window_width}x{window_height}+{x_pos}+{y_pos}")
        self.root.attributes('-topmost', True)
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        self.create_widgets()
        # [추가됨] 시작하자마자 백그라운드에서 업데이트 체크 실행
        threading.Thread(target=self.check_update, daemon=True).start()

    # [새로 추가됨] 자동 업데이트 확인 및 실행 로직
    def check_update(self):
        try:
            requests.packages.urllib3.disable_warnings()
            
            # 1. 서버 버전 확인
            response = requests.get(REPO_URL_VERSION, timeout=3, verify=False)
            
            if response.status_code == 200:
                server_version = response.text.strip()
                
                # 2. 버전 비교 (서버 버전이 더 높으면)
                if server_version > CURRENT_VERSION:
                    self.log(f"🔔 새 버전 발견! ({CURRENT_VERSION} -> {server_version})")
                    
                    if messagebox.askyesno("업데이트", f"새로운 버전({server_version})이 있습니다.\n자동으로 설치를 진행하시겠습니까?\n(프로그램이 재시작됩니다)"):
                        self.log("📥 설치 파일 다운로드 중... 잠시만 기다려주세요.")
                        
                        # 3. 설치 파일 다운로드
                        exe_response = requests.get(REPO_URL_INSTALLER, verify=False)
                        if exe_response.status_code == 200:
                            temp_dir = os.getenv("TEMP")
                            installer_path = os.path.join(temp_dir, f"AutoWork_Update_{server_version}.exe")
                            
                            with open(installer_path, "wb") as f:
                                f.write(exe_response.content)
                            
                            self.log("✅ 다운로드 완료. 업데이트를 시작합니다.")
                            
                            # 4. 설치 파일 실행 (/S 옵션으로 무인 설치)
                            cmd = f'"{installer_path}" /S'
                            subprocess.Popen(cmd, shell=True)
                            
                            # 5. 내 프로그램 즉시 종료 (바통 터치)
                            self.root.destroy()
                            sys.exit(0)
                        else:
                            self.log("❌ 설치 파일 다운로드 실패")
                            messagebox.showerror("에러", "업데이트 파일을 받아오지 못했습니다.")
            
        except Exception as e:
            # 실패 시 조용히 넘어감 (로그만 남김)
            print(f"업데이트 체크 패스: {e}")

    def on_closing(self):
            try:
                # [수정됨] 기다리지 않고(비동기) 백그라운드에서 죽이라고 명령만 던지고 바로 종료
                # os.system 대신 subprocess.Popen 사용 -> 딜레이 0초
                subprocess.Popen("taskkill /f /im chromedriver.exe", shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            except:
                pass
                
            self.root.destroy()
            os._exit(0)

    def create_widgets(self):
        main_paned = tk.PanedWindow(self.root, orient="vertical")
        main_paned.pack(fill="both", expand=True, padx=10, pady=10)

        top_frame = tk.Frame(main_paned)
        
        # [수정] 상단 영역 높이를 600으로 늘림 -> 로그창이 상대적으로 작아짐 (나머지 공간 차지)
        main_paned.add(top_frame, height=510)

        # [왼쪽] 영역
        left_col = tk.Frame(top_frame)
        left_col.pack(side="left", fill="both", expand=True, padx=(0, 10))

        # === [디자인 핵심] 모든 요소를 담을 컨테이너 (달력 너비 기준) ===
        # 이 프레임 안에 다 넣어서 너비를 맞춥니다.
        layout_box = tk.Frame(left_col)
        layout_box.pack(anchor="center", pady=5)

        # ▼▼▼ [추가됨] 사용 설명서 버튼 (맨 위) ▼▼▼
        self.btn_help = tk.Button(layout_box, text="📖 프로그램 사용 설명서 (Click)", bg="#862633", fg="white", font=("bold", 11), command=self.show_guide)
        self.btn_help.pack(fill="x", pady=2)
        # ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲


        # [추가] 크롬 열기 버튼
        self.btn_open_chrome = tk.Button(layout_box, text="1. 🌐 특근입력용 크롬 열기", bg="#34495e", fg="white", font=("bold", 11), height=1, command=self.open_debug_chrome)
        self.btn_open_chrome.pack(fill="x", pady=(10, 5))

        # 1. 달력
        tk.Label(layout_box, text="2. 날짜 선택", font=("맑은 고딕", 11, "bold")).pack(anchor="w")
        self.cal = SimpleCalendar(layout_box, self.update_listbox)
        self.cal.pack(pady=0) # anchor 안 씀 (기본이 center)

        # 2. 시간 설정
        tk.Label(layout_box, text="3. 시간 설정", font=("맑은 고딕", 11, "bold")).pack(anchor="w", pady=(5, 0))
        
        # fill="x"를 사용하여 layout_box(달력너비) 만큼 꽉 채움
        time_frame = tk.Frame(layout_box, bg="#f0f0f0", bd=0)
        time_frame.pack(fill="x", ipady=2) 

        # 내부 요소들 (가운데 정렬을 위해 inner_frame 사용)
        tf_inner = tk.Frame(time_frame, bg="#f0f0f0")
        tf_inner.pack(anchor="center")

        tk.Label(tf_inner, text="시작", bg="#f0f0f0", font=("맑은 고딕", 9)).pack(side="left", padx=(5, 2))
        self.start_h_var = tk.StringVar(value="07")
        self.combo_start_h = ttk.Combobox(tf_inner, textvariable=self.start_h_var, values=[f"{i:02d}" for i in range(24)], width=2)
        self.combo_start_h.pack(side="left")
        tk.Label(tf_inner, text=":", bg="#f0f0f0").pack(side="left")
        
        self.start_m_var = tk.StringVar(value="30")
        self.combo_start_m = ttk.Combobox(tf_inner, textvariable=self.start_m_var, values=["00", "30"], width=2, state="readonly")
        self.combo_start_m.pack(side="left")

        tk.Label(tf_inner, text="~", bg="#f0f0f0", font=("bold")).pack(side="left", padx=2)

        tk.Label(tf_inner, text="종료", bg="#f0f0f0", font=("맑은 고딕", 9)).pack(side="left", padx=(2, 5))

        self.end_h_var = tk.StringVar(value="08")
        self.combo_end_h = ttk.Combobox(tf_inner, textvariable=self.end_h_var, values=[f"{i:02d}" for i in range(24)], width=2)
        self.combo_end_h.pack(side="left")
        tk.Label(tf_inner, text=":", bg="#f0f0f0").pack(side="left")
        
        self.end_m_var = tk.StringVar(value="30")
        self.combo_end_m = ttk.Combobox(tf_inner, textvariable=self.end_m_var, values=["00", "30"], width=2, state="readonly")
        self.combo_end_m.pack(side="left")


        # 3. 사유 선택
        tk.Label(layout_box, text="4. 사유 선택", font=("맑은 고딕", 11, "bold")).pack(anchor="w", pady=(5, 0))
        
        # fill="x"로 달력 너비에 맞춤
        self.reason_var = tk.StringVar(value="조기출근")
        # [수정] 목록에 "직접입력" 추가
        reason_values = ["조기출근", "병동채혈 후 자동화검사실 근무", "병동채혈 후 자동화면역검사실 근무","업무량 증가로 인한 연장근무", "직접입력"]
        
        self.combo_reason = ttk.Combobox(layout_box, textvariable=self.reason_var, values=reason_values, font=("맑은 고딕", 10))
        self.combo_reason.pack(fill="x")
        
        # [추가] 콤보박스 선택 시 실행할 함수 연결
        self.combo_reason.bind("<<ComboboxSelected>>", self.on_reason_select)

        # 4. 실행 버튼
        # fill="x"로 달력 너비에 맞춤
        self.btn_run = tk.Button(layout_box, text="5. 자동 입력 시작", bg="#27ae60", fg="white", font=("bold", 12), height=1, command=self.start_thread)
        self.btn_run.pack(fill="x", pady=5)

        tk.Label(layout_box, text="작업이 끝나면 반드시 신청내역 조회에서\n제대로 입력됐는지 확인하세요.", font=("맑은 고딕", 11, "bold"), fg="red", justify="center").pack(anchor="w")
        self.cal.pack(pady=0) # anchor 안 씀 (기본이 center)


        # [오른쪽] 목록
        right_col = tk.LabelFrame(top_frame, text="선택 목록", font=("맑은 고딕", 10))
        right_col.pack(side="right", fill="x", expand=True, anchor="n", padx=(0, 5))

        self.listbox = tk.Listbox(right_col, font=("consolas", 12), bg="#f9f9f9", height=22)
        self.listbox.pack(side="top", fill="x", padx=5, pady=5)
        ttk.Button(right_col, text="목록 지우기", command=self.clear_dates).pack(side="top", fill="x", padx=5, pady=2)

        # [하단] 로그
        log_frame = tk.LabelFrame(main_paned, text="진행 로그")
        main_paned.add(log_frame)
        self.log_area = scrolledtext.ScrolledText(log_frame, state='disabled', bg='#222222', fg='#00ff00', font=("맑은 고딕", 9))
        self.log_area.pack(fill="both", expand=True)

    # [새로 추가된 함수] 사용 설명서 팝업창
    def show_guide(self):
        guide_window = tk.Toplevel(self.root)
        guide_window.title("프로그램 사용법")
        
        # [위치 계산] 메인 창의 현재 위치(x, y)를 가져옴
        main_x = self.root.winfo_x()
        main_y = self.root.winfo_y()   

        # 메인 창 왼쪽으로 500px 이동 (너비가 500이므로)
        guide_x = main_x - 505  # 5px 간격 둠
        if guide_x < 0: guide_x = 0 # 화면 밖으로 나가면 0으로 고정
        
        # 크기 및 위치 설정 (500x600 사이즈, 위치는 계산된 값)
        guide_window.geometry(f"500x600+{guide_x}+{main_y}")           
        # 설명 텍스트
        info_text = """
[ 1단계: 준비 ]
1. '특근입력용 크롬 열기' 버튼을 누르세요.
2. 새로 열린 크롬창에서 그룹웨어에 로그인하세요.
3. HRM에 들어가서 근태-특근신청 창을 여세요.  

[ 2단계: 설정 ]
1. 달력에서 날짜를 클릭하세요. (파란색 = 선택됨)
   - 잘못 누른 날짜는 다시 누르면 취소됩니다.
2. 시작/종료 시간을 설정하세요.
3. 사유를 선택하세요.
   - '직접입력'을 선택하면 내용을 직접 타이핑할 수 있습니다.

[ 3단계: 실행 ]
1. 초록색 '자동 입력 시작' 버튼을 누르세요.
2. 프로그램이 마우스를 제어하므로 작업이 끝날 때까지 기다려주세요.
3. 작업이 끝나면 **반드시** 신청내역 조회에서 
   제대로 입력됐는지 확인하세요.

[ 기타 기능 ]
■ 선택 목록: 우측 목록에서 내가 선택한 날짜들을 확인할 수 있습니다.
■ 목록 지우기: 선택한 모든 날짜를 초기화합니다.
■ 진행 로그:
   - ✅ 성공: 정상적으로 입력됨
   - ⚠️ 중복: 이미 상신된 내역이 있음 (건너뜀)
   - ⛔ 오류: 미래일시 등 상신 불가 사유 발생
        """
        
        lbl = tk.Label(guide_window, text=info_text, justify="left", font=("맑은 고딕", 11), padx=20, pady=20)
        lbl.pack(fill="both", expand=True)
        
        btn_close = tk.Button(guide_window, text="닫기", command=guide_window.destroy, bg="#34495e", fg="white")
        btn_close.pack(pady=10, ipadx=20)

    # [새로 추가된 함수] 사유 목록에서 선택했을 때 작동
    def on_reason_select(self, event):
        if self.combo_reason.get() == "직접입력":
            self.combo_reason.set("")     # 칸을 비워줌
            self.combo_reason.focus_set() # 바로 타이핑할 수 있게 커서 둠

    # [수정됨] EXE가 있는 폴더(외부 파일)를 찾도록 수정
    def resource_path(self, relative_path):
        try:
            # PyInstaller로 빌드된 경우, 실행 파일(exe)이 있는 실제 폴더를 기준
            if getattr(sys, 'frozen', False):
                base_path = os.path.dirname(sys.executable)
            else:
                base_path = os.path.abspath(".")
        except Exception:
            base_path = os.path.abspath(".")

        return os.path.join(base_path, relative_path)

    # [추가된 함수] 디버깅용 크롬 실행
    def open_debug_chrome(self):
        try:
            # 크롬 경로 찾기
            chrome_paths = [
                r"C:\Program Files\Google\Chrome\Application\chrome.exe",
                r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
            ]
            chrome_path = None
            for path in chrome_paths:
                if os.path.exists(path):
                    chrome_path = path
                    break
            
            if not chrome_path:
                self.log("❌ PC에 설치된 크롬을 찾을 수 없습니다.")
                messagebox.showerror("에러", "크롬이 설치된 경로를 찾을 수 없습니다.")
                return

            # AppData/Local 폴더에 프로필 생성
            app_data_path = os.getenv('LOCALAPPDATA') 
            user_data_dir = os.path.join(app_data_path, "AutoWork_Chrome_Profile")
            
            # 폴더가 없으면 생성 (경로상의 모든 폴더 자동 생성)
            if not os.path.exists(user_data_dir):
                os.makedirs(user_data_dir)

            # 3. [핵심] 화면 크기 계산하여 우측 상단 좌표 구하기
            screen_width = self.root.winfo_screenwidth() # 내 모니터 가로 크기
            target_w = 1290
            target_h = 1030
            
            # 우측 끝에 붙이려면: (전체 너비 - 창 너비)가 X좌표가 됨
            pos_x = 0
            pos_y = 0  # 상단은 0
            

            # 4. 실행 명령어 (사이즈 및 위치 옵션 추가)
            # --window-size=너비,높이
            # --window-position=X,Y
            cmd = (f'"{chrome_path}" '
                   f'--remote-debugging-port=9222 '
                   f'--user-data-dir="{user_data_dir}" '
                   f'--window-size={target_w},{target_h} '
                   f'--window-position={pos_x},{pos_y} '
                   f'"https://gw.kumc.or.kr/"')
            
            # 비동기로 실행 (파이썬이 멈추지 않게 Popen 사용)
            subprocess.Popen(cmd, shell=True)
            self.log("🚀 특근입력용 크롬을 실행했습니다.")
            self.log("⚠️ 새로 열린 크롬에서 로그인을 먼저 해주세요!")

        except Exception as e:
            self.log(f"❌ 크롬 실행 실패: {e}")
            messagebox.showerror("에러", f"크롬 실행 중 오류가 발생했습니다.\n{e}")

    # # [교체됨] PC 설치 크롬 대신, 내장된 'Chrome for Testing'을 실행
    # def open_debug_chrome(self):
    #     try:
    #         # 1. 내장된 특수 크롬(chrome-win64/chrome.exe) 경로 찾기
    #         local_chrome_path = self.resource_path(os.path.join("chrome-win64", "chrome.exe"))
            
    #         if os.path.exists(local_chrome_path):
    #             chrome_path = local_chrome_path
    #             self.log(f"🔧 내장된 특수 크롬을 사용합니다.\n경로: {chrome_path}")
    #         else:
    #             self.log("❌ 내장 크롬(chrome-win64)을 찾을 수 없습니다.")
    #             messagebox.showerror("파일 없음", 
    #                 "프로그램 폴더 안에 'chrome-win64' 폴더가 없습니다.\n"
    #                 "Chrome for Testing 파일을 다운받아 넣어주세요.")
    #             return

    #         # 2. 프로필 폴더 설정 (기존과 동일하지만 폴더명 변경 권장)
    #         app_data_path = os.getenv('LOCALAPPDATA') 
    #         user_data_dir = os.path.join(app_data_path, "AutoWork_Chrome_Profile_Fixed")
            
    #         if not os.path.exists(user_data_dir):
    #             os.makedirs(user_data_dir)

    #         # 3. 화면 크기 및 위치
    #         target_w = 1290
    #         target_h = 1030
    #         pos_x = 0
    #         pos_y = 0

    #         # 4. 실행 명령어
    #         cmd = (f'"{chrome_path}" '
    #                f'--remote-debugging-port=9222 '
    #                f'--user-data-dir="{user_data_dir}" '
    #                f'--window-size={target_w},{target_h} '
    #                f'--window-position={pos_x},{pos_y} '
    #                f'"https://gw.kumc.or.kr/"')

    #         subprocess.Popen(cmd, shell=True)
    #         self.log(f"🚀 특수 크롬(v고정) 실행 완료!")
    #         self.log("⚠️ 새로 열린 크롬에서 로그인을 먼저 해주세요!")

    #     except Exception as e:
    #         self.log(f"❌ 크롬 실행 실패: {e}")
    #         messagebox.showerror("에러", f"크롬 실행 중 오류가 발생했습니다.\n{e}")


    def update_listbox(self, selected_dates):
        self.listbox.delete(0, tk.END)
        for d in sorted(list(selected_dates)):
            self.listbox.insert(tk.END, d)

    def clear_dates(self):
        self.cal.selected_dates.clear()
        self.update_listbox(set())
        self.cal.draw_days()

    def log(self, msg):
        self.log_area.config(state='normal')
        self.log_area.insert(tk.END, f"[{datetime.now().strftime('%H:%M:%S')}] {msg}\n")
        self.log_area.see(tk.END)
        self.log_area.config(state='disabled')

    def find_element_recursive(self, driver, xpath):
        try: return driver.find_element(By.XPATH, xpath)
        except: pass
        frames = driver.find_elements(By.TAG_NAME, "frame") + driver.find_elements(By.TAG_NAME, "iframe")
        for frame in frames:
            try:
                driver.switch_to.frame(frame)
                found = self.find_element_recursive(driver, xpath)
                if found: return found
                driver.switch_to.parent_frame()
            except: driver.switch_to.parent_frame()
        return None

    def start_thread(self):
        if not self.cal.selected_dates:
            messagebox.showwarning("알림", "날짜를 선택하세요.")
            return
        self.btn_run.config(state='disabled')
        threading.Thread(target=self.run_macro, daemon=True).start()

    def run_macro(self):
        # [핵심 추가] 0. 디버깅 크롬이 켜져있는지 먼저 확인 (포트 9222 체크)
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        sock.settimeout(1) # 1초 안에 응답 없으면 닫힘
        result = sock.connect_ex(('127.0.0.1', 9222))
        sock.close()

        if result != 0:
            # 0이 아니면 연결 실패 (크롬이 안 켜져 있음)
            msg = "특근입력용 크롬이 열려있지 않습니다.\n파란색 버튼을 눌러 크롬을 열고\n그룹웨어 로그인－HRM 접속－근태－특근신청 누르고\n다시 시도하세요"
            self.log("⛔ " + msg.replace("\n", " ")) # 로그에는 한 줄로
            messagebox.showwarning("실행 불가", msg) # 팝업
            self.btn_run.config(state='normal')
            return


        # ▼▼▼ 드라이버 자동 맞춤 로직 실행 ▼▼▼
        updater = ChromeDriverUpdater()
        driver_path = self.resource_path("chromedriver.exe") # 경로 설정
        
        # 드라이버 업데이트 시도 (실패하면 매크로 중단)
        # resource_path로 감싸진 경로는 읽기 전용일 수 있으므로, 
        # 실제 실행 파일이 있는 폴더에 다운로드하도록 로직 조정 필요.
        # 여기서는 간단히 실행 위치의 chromedriver.exe를 체크합니다.
        
        if not updater.update_driver_if_needed("chromedriver.exe"):
            self.log("⛔ 드라이버 준비 실패로 작업을 중단합니다.")
            self.btn_run.config(state='normal')
            return
        # ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

            # ... (이후 기존 셀레니움 실행 코드) ...
            # driver = webdriver.Chrome(service=Service("chromedriver.exe"), ...) 
            # 주의: Service 경로를 업데이트된 파일인 "chromedriver.exe"로 지정해야 함

        # ---------------------------------------------------------
        # 이하 기존 로직 (크롬이 켜져 있을 때만 실행됨)
        # ---------------------------------------------------------
        try:
            chrome_options = Options()
            chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
            driver = None        

            try:
                self.log("🔧 내장 드라이버 연결 시도...")
                driver_path = self.resource_path("chromedriver.exe")
                
                # 파일 존재 여부 확인
                if os.path.exists(driver_path):
                    service = Service(executable_path=driver_path)
                    driver = webdriver.Chrome(service=service, options=chrome_options)
                    self.log("✅ 내장 드라이버 연결 성공")
                else:
                    raise FileNotFoundError("로컬 파일 없음")
                    
            except Exception:
                try:
                    self.log("⚠️ 내장 드라이버 없음/실패 -> 자동 다운로드 시도...")
                    from webdriver_manager.chrome import ChromeDriverManager
                    driver_path = ChromeDriverManager().install()
                    service = Service(executable_path=driver_path)
                    driver = webdriver.Chrome(service=service, options=chrome_options)
                    self.log("✅ 자동 다운로드 드라이버 연결 성공")
                except Exception as e2:
                    self.log(f"❌ 드라이버 연결 실패: {e2}")
                    messagebox.showerror("에러", f"드라이버 연결 불가.\n인터넷 연결 또는 chromedriver.exe 파일을 확인하세요.\n{e2}")
                    return

            # [창 찾기 로직]
            main_window = None
            for handle in driver.window_handles:
                driver.switch_to.window(handle)
                driver.switch_to.default_content()
                if self.find_element_recursive(driver, XPATHS["CALENDAR_ICON"]):
                    main_window = handle
                    self.log("✅ 특근 창을 확인했습니다.")
                    break
            
            if not main_window:
                for handle in driver.window_handles:
                    driver.switch_to.window(handle)
                    driver.switch_to.default_content()
                    if self.find_element_recursive(driver, XPATHS["SUBMIT_BTN"]):
                        main_window = handle
                        self.log("✅ 특근 창을 확인했습니다. (버튼 기준)")
                        break

            if not main_window:
                self.log("❌ 특근신청 창을 못 찾았습니다.")
                self.btn_run.config(state='normal')
                return

            driver.switch_to.window(main_window)
            driver.switch_to.default_content()

            target_dates = sorted(list(self.cal.selected_dates))
            
            # [집계 변수]
            success_cnt = 0
            dup_cnt = 0
            error_details = [] 
            duplicate_details = []
            success_details = [] # [추가됨] 성공 내역 저장용 리스트

            for idx, date_str in enumerate(target_dates):
                self.log(f"▶ [{date_str}] 진행...")
                day_num = str(int(date_str.split('-')[2]))

                try:
                    driver.switch_to.window(main_window)
                    driver.switch_to.default_content()

                    # [헬퍼 함수들]
                    def pure_js_click(xpath):
                        driver.switch_to.default_content()
                        elem = self.find_element_recursive(driver, xpath)
                        if elem:
                            driver.execute_script("arguments[0].click();", elem)
                            return True
                        return False

                    def pure_js_inject(xpath, value):
                        driver.switch_to.default_content()
                        elem = self.find_element_recursive(driver, xpath)
                        if elem:
                            driver.execute_script("arguments[0].value = '';", elem)
                            driver.execute_script(f"arguments[0].value = '{value}';", elem)
                            driver.execute_script("arguments[0].dispatchEvent(new Event('change'));", elem)
                            return True
                        return False

                    # (A) 달력 아이콘
                    if not pure_js_click(XPATHS["CALENDAR_ICON"]):
                        self.log("❌ 달력 아이콘 실패")
                        continue
                    time.sleep(0.2)

                    # (B) 날짜 클릭
                    found_date = False
                    for _ in range(3):
                        driver.switch_to.default_content()
                        try:
                            try: date_el = driver.find_element(By.LINK_TEXT, day_num)
                            except: date_el = self.find_element_recursive(driver, f"//*[text()='{day_num}']")
                            if date_el:
                                driver.execute_script("arguments[0].click();", date_el)
                                found_date = True
                                break
                        except: pass
                        time.sleep(0.2)
                    if not found_date:
                        self.log(f"❌ 날짜 [{day_num}] 클릭 실패")
                        continue
                    time.sleep(0.2)

                    # (C) 시간 입력
                    if idx == 0:
                        self.log("   (시간 설정...)")
                        pure_js_click(XPATHS["START_H_BTN"])
                        time.sleep(0.3)
                        item_xpath = f"//*[@id='cmbSTRT_HHXX_itemTable_{int(self.start_h_var.get())}']"
                        if not pure_js_click(item_xpath):
                            try:
                                el = driver.find_element(By.XPATH, item_xpath)
                                driver.execute_script("arguments[0].click();", el)
                            except: pass
                        time.sleep(0.1)
                        pure_js_inject(XPATHS["START_M_IPT"], self.start_m_var.get())

                        pure_js_click(XPATHS["END_H_BTN"])
                        time.sleep(0.3)
                        item_xpath = f"//*[@id='cmbENDX_HHXX_itemTable_{int(self.end_h_var.get())}']"
                        if not pure_js_click(item_xpath):
                            try:
                                el = driver.find_element(By.XPATH, item_xpath)
                                driver.execute_script("arguments[0].click();", el)
                            except: pass
                        time.sleep(0.1)
                        pure_js_inject(XPATHS["END_M_IPT"], self.end_m_var.get())
                        time.sleep(0.5)    

                    # (D) 사유 입력
                    target_reason = self.reason_var.get()
                    try:
                        driver.switch_to.default_content()
                        el = self.find_element_recursive(driver, XPATHS["REASON_INPUT"])
                        if el:
                            driver.execute_script("arguments[0].click();", el)
                            time.sleep(0.2)
                            el.send_keys(Keys.CONTROL, 'a')
                            el.send_keys(Keys.DELETE)
                            time.sleep(0.1)
                            el.send_keys(target_reason)
                            time.sleep(0.1)
                            el.send_keys(Keys.TAB)
                        else:
                            self.log("⚠️ 사유 입력칸을 못 찾았습니다.")
                    except Exception as e:
                        self.log(f"⚠️ 사유 입력 중 에러: {e}")
                        try:
                            if el: driver.execute_script(f"arguments[0].value = '{target_reason}';", el)
                        except: pass
                    time.sleep(0.5)
                                                        
                    # 상신
                    if not pure_js_click(XPATHS["SUBMIT_BTN"]):
                        self.log("❌ 상신 버튼 실패")
                        continue
                    time.sleep(0.5)

                    # [판독 로직]
                    result_processed = False
                    try:
                        for _ in range(6): 
                            driver.switch_to.default_content()
                            
                            dup_msg = self.find_element_recursive(driver, XPATHS["DUPLICATE_MSG"])
                            if dup_msg and dup_msg.is_displayed():
                                msg_text = dup_msg.text.strip()
                                
                                if "정상적으로 상신되었습니다" in msg_text:
                                    self.log(f"✅ [{date_str}] 성공!")
                                    success_cnt += 1
                                    success_details.append(f"[{date_str}] 성공") # [추가]
                                elif "중복" in msg_text or "이미" in msg_text:
                                    short_msg = msg_text.split('\n')[0]
                                    self.log(f"⚠️ [{date_str}] 중복: {short_msg}")
                                    dup_cnt += 1
                                    duplicate_details.append(f"[{date_str}] {short_msg}")
                                else:
                                    short_msg = msg_text.split('\n')[0]
                                    self.log(f"⛔ [{date_str}] 실패: {short_msg}...")
                                    error_details.append(f"[{date_str}] {short_msg}")
                                
                                pure_js_click(XPATHS["POPUP_CONFIRM"])
                                result_processed = True
                                break
                            # 2. 버튼 확인 (성공)
                            confirm_btn = self.find_element_recursive(driver, XPATHS["POPUP_CONFIRM"])
                            if confirm_btn and confirm_btn.is_displayed():
                                driver.execute_script("arguments[0].click();", confirm_btn)
                                self.log(f"✅ [{date_str}] 성공 (버튼확인)!")
                                success_cnt += 1
                                success_details.append(f"[{date_str}] 성공") # [추가]
                                result_processed = True
                                break
                            # 3. 브라우저 Alert 확인
                            try:
                                alert = driver.switch_to.alert
                                alert_text = alert.text
                                if "정상적으로" in alert_text:
                                    self.log(f"✅ [{date_str}] 성공!")
                                    success_cnt += 1
                                    success_details.append(f"[{date_str}] 성공") # [추가]
                                elif "중복" in alert_text or "이미" in alert_text:
                                    self.log(f"⚠️ [{date_str}] 중복: {alert_text}")
                                    dup_cnt += 1
                                    duplicate_details.append(f"[{date_str}] {alert_text}")
                                else:
                                    self.log(f"⛔ [{date_str}] 실패: {alert_text}")
                                    error_details.append(f"[{date_str}] {alert_text}")
                                alert.accept()
                                result_processed = True
                                break
                            except: pass
                            
                            time.sleep(0.5)
                                
                    except Exception as e:
                        self.log(f"   ⚠️ 팝업 판독 오류: {e}")

                    if not result_processed:
                        self.log(f"❌ [{date_str}] 실패 (팝업 안 뜸/응답 없음)")
                        error_details.append(f"[{date_str}] 응답 없음")

                    time.sleep(0.2)

                except Exception as e:
                    self.log(f"❌ 에러: {e}")

            self.log("🎉 작업 종료")
            # [수정됨] 최종 리포트 함수 호출 (성공 내역 리스트도 전달)
            self._show_final_report(success_cnt, dup_cnt, error_details, duplicate_details, success_details)
            
        except Exception as e:
            self.log(f"연결 실패: {e}")
        finally:
            self.btn_run.config(state='normal')
    
    # [수정됨] 최종 결과 보고서 (성공/중복/오류 모든 상세 내역 포함)
    def _show_final_report(self, success_cnt, dup_cnt, error_details, duplicate_details, success_details):
        total_cnt = success_cnt + dup_cnt + len(error_details)
        error_cnt = len(error_details)
        
        # 1. 요약 메시지
        summary_text = (
            f"📊 [최종 결과 요약]\n"
            f"총 {total_cnt}건 처리 완료\n"
            f"----------------------------\n"
            f"✅ 성공 : {success_cnt}건\n"
            f"⚠️ 중복 : {dup_cnt}건\n"
            f"⛔ 오류 : {error_cnt}건"
        )
        
        # 2. 상세 내역 텍스트 생성
        details_text = ""
        
        # [추가됨] 성공 내역 표시
        if success_details:
            details_text += "\n\n[✅ 성공한 항목]\n" + "\n".join(success_details)

        if duplicate_details:
            details_text += "\n\n[⚠️ 중복된 항목]\n" + "\n".join(duplicate_details)
            
        if error_details:
            details_text += "\n\n[⛔ 오류 발생 항목]\n" + "\n".join(error_details)

        # 3. 로그창 출력
        self.log("\n" + "="*35)
        self.log(summary_text)
        if details_text:
            self.log(details_text) # 상세 내용 로그에 출력
        self.log("\n" + "="*35 + "\n")
        
        # 4. 팝업창 출력
        full_msg = summary_text + details_text
        
        if error_details or duplicate_details:
            # 확인이 필요한 사항이 있으면 Warning
            messagebox.showwarning("작업 완료 (확인 필요)", full_msg)
        else:
            # 모두 성공했으면 Info
            messagebox.showinfo("작업 완료", full_msg)

if __name__ == "__main__":
    root = tk.Tk()
    app = AutoWorkApp(root)
    root.mainloop()

# pyinstaller -w -F --icon=jjangu3.ico --exclude-module pandas --exclude-module numpy --exclude-module PIL --add-binary "chromedriver.exe;." Special_Work_Writer8.py