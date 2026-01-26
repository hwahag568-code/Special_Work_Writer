import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import threading
import time
from datetime import datetime
import calendar
import os
import sys
import subprocess
import socket
import requests
import winreg
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import urllib3

# 보안 경고 무시
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ==============================================================================
# [설정] 사용자 환경 및 GitHub 정보 (수정 필요)
# ==============================================================================
CURRENT_VERSION = "1.5"

# 버전 파일 및 설치 파일 주소
REPO_URL_VERSION = "https://raw.githubusercontent.com/hwahag568-code/Special_Work_Writer/main/version.txt"
REPO_URL_INSTALLER = "https://raw.githubusercontent.com/hwahag568-code/Special_Work_Writer/main/Update_Work_Writer.exe"

# 드라이버 다운로드 주소 (drivers 폴더 경로 포함)
DRIVER_BASE_URL = "https://raw.githubusercontent.com/hwahag568-code/Special_Work_Writer/main/drivers/chromedriver_"

# [XPath 설정]
XPATHS = {
    "CALENDAR_ICON": '//*[@id="ipcSTRT_YMDX_img"]',
    "REASON_INPUT":  '//*[@id="iptAPPL_RMRK"]',
    "SUBMIT_BTN":    '//*[@id="btnSetEMPL_NUMB_center"]',
    "START_H_BTN":   '//*[@id="cmbSTRT_HHXX_button"]',
    "START_M_IPT":   '//*[@id="iptSTRT_MMXX"]',
    "END_H_BTN":     '//*[@id="cmbENDX_HHXX_button"]',
    "END_M_IPT":     '//*[@id="iptENDX_MMXX"]',
    "POPUP_CONFIRM": '//*[@id="btn_confirm"]', 
    "DUPLICATE_MSG": '//*[@id="grpMessage"]',
    # 로그인 여부 판단용 (메인화면 바로가기 버튼)
    "MAIN_SHORTCUT": '//*[@id="util_quickLink"]' 
}

# ==============================================================================
# [클래스 1] 달력 위젯 (변경 없음)
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
        header = tk.Frame(self)
        header.pack(fill="x", pady=2)
        tk.Button(header, text="<", command=self.prev_month).pack(side="left")
        self.lbl_header = tk.Label(header, text=f"{self.year}년 {self.month}월", font=("Arial", 11, "bold"))
        self.lbl_header.pack(side="left", expand=True)
        tk.Button(header, text=">", command=self.next_month).pack(side="right")
        self.grid_frame = tk.Frame(self)
        self.grid_frame.pack()
        self.draw_days()

    def draw_days(self):
        for widget in self.grid_frame.winfo_children():
            widget.destroy()
        self.buttons = {}
        days = ["일", "월", "화", "수", "목", "금", "토"]
        for i, d in enumerate(days):
            lbl = tk.Label(self.grid_frame, text=d, width=4, fg="red" if d=="일" else "black")
            lbl.grid(row=0, column=i, padx=1, pady=(0, 5))
        cal = calendar.Calendar(firstweekday=6).monthdayscalendar(self.year, self.month)
        for r, week in enumerate(cal):
            for c, day in enumerate(week):
                if day != 0:
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

# ==============================================================================
# [클래스 2] 드라이버 업데이트 관리자
# ==============================================================================
class ChromeDriverUpdater:
    def get_chrome_version(self):
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Google\Chrome\BLBeacon")
            version, _ = winreg.QueryValueEx(key, "version")
            return version.split('.')[0]
        except:
            try:
                key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Google\Chrome\BLBeacon")
                version, _ = winreg.QueryValueEx(key, "version")
                return version.split('.')[0]
            except:
                return None

    def get_driver_version(self, driver_path):
        if not os.path.exists(driver_path): return None
        try:
            result = subprocess.check_output([driver_path, "--version"], stderr=subprocess.STDOUT)
            return result.decode('utf-8').split(' ')[1].split('.')[0]
        except: return None

    def update_driver_if_needed(self, driver_name="chromedriver.exe"):
        chrome_ver = self.get_chrome_version()
        if not chrome_ver:
            print("❌ 크롬 버전을 찾을 수 없습니다.")
            return False

        driver_ver = self.get_driver_version(driver_name)
        if chrome_ver == driver_ver:
            return True

        # 다운로드 (내 컴퓨터 크롬 버전에 맞는 것만 쏙 골라오기)
        download_url = f"{DRIVER_BASE_URL}{chrome_ver}.exe"
        try:
            print(f"🔄 드라이버 업데이트 시도 (v{chrome_ver})...")
            response = requests.get(download_url, verify=False, timeout=10)
            if response.status_code == 200:
                if os.path.exists(driver_name):
                    try:
                        os.remove(driver_name)
                    except:
                        subprocess.call(f"taskkill /f /im {driver_name}", shell=True)
                        time.sleep(1)
                        os.remove(driver_name)
                
                with open(driver_name, "wb") as f:
                    f.write(response.content)
                print("🎉 드라이버 업데이트 완료")
                return True
            else:
                messagebox.showerror("오류", f"서버에 v{chrome_ver} 드라이버가 없습니다.\n관리자 도구를 실행해 버전을 추가해주세요.")
                return False
        except Exception as e:
            messagebox.showerror("오류", f"드라이버 다운로드 실패: {e}")
            return False

# ==============================================================================
# [클래스 3] 메인 어플리케이션
# ==============================================================================
class AutoWorkApp:
    def __init__(self, root):
        self.root = root
        self.root.title(f"특근 입력기 v{CURRENT_VERSION}")
        
        window_width = 430
        window_height = 700 # 높이를 조금 더 늘림 (로그인창 추가됨)
        screen_width = root.winfo_screenwidth()
        x_pos = screen_width - window_width - 200
        y_pos = 100
        self.root.geometry(f"{window_width}x{window_height}+{x_pos}+{y_pos}")
        self.root.attributes('-topmost', True)
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        self.create_widgets()
        threading.Thread(target=self.check_update, daemon=True).start()

    def check_update(self):
        try:
            response = requests.get(REPO_URL_VERSION, timeout=3, verify=False)
            if response.status_code == 200:
                server_version = response.text.strip()
                if server_version > CURRENT_VERSION:
                    self.log(f"🔔 업데이트 발견 ({CURRENT_VERSION} -> {server_version})")
                    if messagebox.askyesno("업데이트", f"새 버전({server_version})이 있습니다.\n업데이트 하시겠습니까?"):
                        exe_response = requests.get(REPO_URL_INSTALLER, verify=False)
                        if exe_response.status_code == 200:
                            temp_dir = os.getenv("TEMP")
                            installer_path = os.path.join(temp_dir, f"AutoWork_Update.exe")
                            with open(installer_path, "wb") as f:
                                f.write(exe_response.content)
                            subprocess.Popen(f'"{installer_path}" /S', shell=True)
                            self.root.destroy()
                            sys.exit(0)
        except: pass

    def on_closing(self):
        subprocess.Popen("taskkill /f /im chromedriver.exe", shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        self.root.destroy()
        os._exit(0)

    def resource_path(self, relative_path):
        try:
            if getattr(sys, 'frozen', False): base_path = os.path.dirname(sys.executable)
            else: base_path = os.path.abspath(".")
        except: base_path = os.path.abspath(".")
        return os.path.join(base_path, relative_path)

    def create_widgets(self):
        main_paned = tk.PanedWindow(self.root, orient="vertical")
        main_paned.pack(fill="both", expand=True, padx=10, pady=10)
        
        top_frame = tk.Frame(main_paned)
        main_paned.add(top_frame, height=560) # 높이 조절

        left_col = tk.Frame(top_frame)
        left_col.pack(side="left", fill="both", expand=True, padx=(0, 10))
        layout_box = tk.Frame(left_col)
        layout_box.pack(anchor="center", pady=5)

        # 0. 사용 설명서
        self.btn_help = tk.Button(layout_box, text="📖 사용 설명서", bg="#862633", fg="white", font=("bold", 10), command=self.show_guide)
        self.btn_help.pack(fill="x", pady=2)

        # 1. 로그인 정보 (추가됨)
        input_frame = tk.LabelFrame(layout_box, text="1. 로그인 정보(그룹웨어)", font=("맑은 고딕", 10, "bold"))
        input_frame.pack(fill="x", pady=(5, 5))
        
        row_frame = tk.Frame(input_frame)
        row_frame.pack(pady=2)
        tk.Label(row_frame, text="ID:").pack(side="left")
        self.entry_id = tk.Entry(row_frame, width=12)
        self.entry_id.pack(side="left", padx=2)
        self.entry_id.insert(0, "12A27") # 기본값
        
        tk.Label(row_frame, text="PW:").pack(side="left", padx=(5,0))
        self.entry_pw = tk.Entry(row_frame, width=12, show="*")
        self.entry_pw.pack(side="left", padx=2)
        self.entry_pw.insert(0, "1") # 기본값


        # 2. 크롬 열기 버튼 (수동용)
        self.btn_open_chrome = tk.Button(layout_box, text="🌐 수동 크롬 열기 (필요시)", bg="#34495e", fg="white", height=1, command=self.open_debug_chrome)
        self.btn_open_chrome.pack(fill="x", pady=(2, 5))

        # 3. 달력
        tk.Label(layout_box, text="2. 날짜 선택", font=("맑은 고딕", 11, "bold")).pack(anchor="w")
        self.cal = SimpleCalendar(layout_box, self.update_listbox)
        self.cal.pack()

        # 4. 시간
        tk.Label(layout_box, text="3. 시간 설정", font=("맑은 고딕", 11, "bold")).pack(anchor="w", pady=(5,0))
        time_frame = tk.Frame(layout_box, bg="#f0f0f0")
        time_frame.pack(fill="x")
        tf_inner = tk.Frame(time_frame, bg="#f0f0f0")
        tf_inner.pack(anchor="center")
        
        tk.Label(tf_inner, text="시작").pack(side="left")
        self.start_h_var = tk.StringVar(value="07")
        ttk.Combobox(tf_inner, textvariable=self.start_h_var, values=[f"{i:02d}" for i in range(24)], width=2).pack(side="left")
        tk.Label(tf_inner, text=":").pack(side="left")
        self.start_m_var = tk.StringVar(value="30")
        ttk.Combobox(tf_inner, textvariable=self.start_m_var, values=["00", "30"], width=2, state="readonly").pack(side="left")
        
        tk.Label(tf_inner, text=" ~ ").pack(side="left")
        
        tk.Label(tf_inner, text="종료").pack(side="left")
        self.end_h_var = tk.StringVar(value="08")
        ttk.Combobox(tf_inner, textvariable=self.end_h_var, values=[f"{i:02d}" for i in range(24)], width=2).pack(side="left")
        tk.Label(tf_inner, text=":").pack(side="left")
        self.end_m_var = tk.StringVar(value="30")
        ttk.Combobox(tf_inner, textvariable=self.end_m_var, values=["00", "30"], width=2, state="readonly").pack(side="left")

        # 5. 사유
        tk.Label(layout_box, text="4. 사유 선택", font=("맑은 고딕", 11, "bold")).pack(anchor="w", pady=(5,0))
        self.reason_var = tk.StringVar(value="조기출근")
        self.combo_reason = ttk.Combobox(layout_box, textvariable=self.reason_var, values=["조기출근", "업무량 증가", "직접입력"])
        self.combo_reason.pack(fill="x")
        self.combo_reason.bind("<<ComboboxSelected>>", lambda e: self.combo_reason.set("") if self.combo_reason.get()=="직접입력" else None)

        # 6. 실행 버튼
        self.btn_run = tk.Button(layout_box, text="▶ 자동 입력 시작", bg="#27ae60", fg="white", font=("bold", 12), command=self.start_thread)
        self.btn_run.pack(fill="x", pady=10)

        # 오른쪽 목록
        right_col = tk.LabelFrame(top_frame, text="선택 목록")
        right_col.pack(side="right", fill="x", expand=True, anchor="n")
        self.listbox = tk.Listbox(right_col, height=20, bg="#f9f9f9")
        self.listbox.pack(fill="x", padx=5, pady=5)
        ttk.Button(right_col, text="목록 지우기", command=self.clear_dates).pack(fill="x", padx=5)

        # 하단 로그
        log_frame = tk.LabelFrame(main_paned, text="로그")
        main_paned.add(log_frame)
        self.log_area = scrolledtext.ScrolledText(log_frame, state='disabled', bg='#222', fg='#0f0', font=("맑은 고딕", 9))
        self.log_area.pack(fill="both", expand=True)

    def log(self, msg):
        self.log_area.config(state='normal')
        self.log_area.insert(tk.END, f"[{datetime.now().strftime('%H:%M:%S')}] {msg}\n")
        self.log_area.see(tk.END)
        self.log_area.config(state='disabled')

    def update_listbox(self, selected_dates):
        self.listbox.delete(0, tk.END)
        for d in sorted(list(selected_dates)):
            self.listbox.insert(tk.END, d)

    def clear_dates(self):
        self.cal.selected_dates.clear()
        self.update_listbox(set())
        self.cal.draw_days()
        
    def show_guide(self):
        messagebox.showinfo("사용법", "1. 날짜 선택\n2. 시간/사유 설정\n3. 시작 버튼 클릭\n\n*로그인이 안 되어 있으면 ID/PW 입력 필수!")

    def open_debug_chrome(self):
        try:
            paths = [r"C:\Program Files\Google\Chrome\Application\chrome.exe", r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"]
            chrome_path = next((p for p in paths if os.path.exists(p)), None)
            if not chrome_path:
                messagebox.showerror("에러", "크롬을 찾을 수 없습니다.")
                return

            app_data = os.getenv('LOCALAPPDATA')
            user_data = os.path.join(app_data, "AutoWork_Profile")
            if not os.path.exists(user_data): os.makedirs(user_data)
            
            # ▼▼▼ [수정됨] --disable-popup-blocking 옵션 추가! ▼▼▼
            cmd = (f'"{chrome_path}" '
                   f'--remote-debugging-port=9222 '
                   f'--user-data-dir="{user_data}" '
                   f'--window-size=1280,1024 '
                   f'--disable-popup-blocking '  # 👈 여기가 핵심입니다 (팝업 차단 해제)
                   f'"https://gw.kumc.or.kr/"')
            
            subprocess.Popen(cmd, shell=True)
            self.log("🚀 크롬 실행 완료 (팝업차단 해제됨)")
        except Exception as e:
            self.log(f"❌ 크롬 실행 에러: {e}")
    
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

    # ==========================================================================
    # [최종 최적화] 스마트 대기 적용된 run_macro
    # ==========================================================================
    def run_macro(self):
        # ---------------------------------------------------------
        # 0. 헬퍼 함수 정의 (스마트 대기용)
        # ---------------------------------------------------------
        # 특정 요소가 클릭 가능해질 때까지 기다렸다가 클릭
        def wait_click(driver, xpath, timeout=10):
            try:
                # 1. 메인 프레임에서 시도
                driver.switch_to.default_content()
                elem = WebDriverWait(driver, 1).until(EC.element_to_be_clickable((By.XPATH, xpath)))
                driver.execute_script("arguments[0].click();", elem)
                return True
            except:
                # 2. 없으면 재귀함수로 찾아서 클릭 (기존 방식 + 즉시 실행)
                elem = self.find_element_recursive(driver, xpath)
                if elem:
                    driver.execute_script("arguments[0].click();", elem)
                    return True
            return False

        # ---------------------------------------------------------
        # 1. 크롬 실행 상태 확인
        # ---------------------------------------------------------
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        sock.settimeout(0.5) # 타임아웃도 줄임 (속도 향상)
        is_chrome_running = (sock.connect_ex(('127.0.0.1', 9222)) == 0)
        sock.close()

        need_login = True
        target_window = None

        # ---------------------------------------------------------
        # 2. 빠른 상태 진단
        # ---------------------------------------------------------
        if is_chrome_running:
            try:
                chrome_options = Options()
                chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
                
                # 드라이버 파일 체크 (생략 가능하지만 안전을 위해 유지)
                if not os.path.exists("chromedriver.exe"):
                     ChromeDriverUpdater().update_driver_if_needed()
                service = Service(executable_path="chromedriver.exe")
                temp_driver = webdriver.Chrome(service=service, options=chrome_options)
                
                # [고속 진단] 현재 열린 모든 창을 빠르게 훑음
                for handle in temp_driver.window_handles:
                    try:
                        temp_driver.switch_to.window(handle)
                        # 프레임 전환 없이 일단 달력 아이콘이 있는지 빠르게 확인 (소스코드 레벨 체크)
                        if "ipcSTRT_YMDX_img" in temp_driver.page_source: 
                            # 소스에 있으면 정밀 확인
                            if self.find_element_recursive(temp_driver, XPATHS['CALENDAR_ICON']):
                                self.log("⚡ 특근 창 발견! 즉시 시작합니다.")
                                target_window = handle
                                need_login = False
                                break
                    except: pass
                
                # 특근창 없으면 로그인 여부 확인
                if need_login:
                    for handle in temp_driver.window_handles:
                        try:
                            temp_driver.switch_to.window(handle)
                            if "util_quickLink" in temp_driver.page_source: # 고속 체크
                                self.log("⚡ 로그인 상태 확인됨.")
                                need_login = False
                                break
                        except: pass
            except: pass

        # ---------------------------------------------------------
        # 3. 로그인 정보 (필요시)
        # ---------------------------------------------------------
        if need_login:
            user_id = self.entry_id.get().strip()
            user_pw = self.entry_pw.get().strip()
            if not user_id or not user_pw:
                messagebox.showwarning("필수", "로그인이 필요합니다.\nID/PW를 입력하세요.")
                self.btn_run.config(state='normal')
                return
            
            if not is_chrome_running:
                self.log("🔄 크롬 실행...")
                self.open_debug_chrome()
                # 켜지자마자 연결 시도하므로 sleep 최소화
                time.sleep(1.5) 

        # ---------------------------------------------------------
        # 4. 드라이버 연결
        # ---------------------------------------------------------
        try:
            chrome_options = Options()
            chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
            chrome_options.add_argument("--disable-popup-blocking") # 팝업 차단 해제
            
            service = Service(executable_path="chromedriver.exe")
            driver = webdriver.Chrome(service=service, options=chrome_options)

            # ==========================================
            # A. 로그인 (스마트 대기 적용)
            # ==========================================
            if need_login:
                self.log("🚀 접속 중...")
                driver.get("https://gw.kumc.or.kr/user/login/login.do")
                
                try:
                    # 입력창이 '보일 때'까지 기다림 (최대 10초) -> 뜨면 0.1초만에 통과
                    wait = WebDriverWait(driver, 10)
                    
                    # 프레임이 있을 수 있으니 재귀 함수로 찾되, 못 찾으면 바로 에러
                    login_input = None
                    for _ in range(5):
                        login_input = self.find_element_recursive(driver, '//*[@id="uid"]')
                        if login_input: break
                        time.sleep(0.5)
                    
                    if login_input:
                        login_input.click(); login_input.clear(); login_input.send_keys(user_id)
                        pw_input = self.find_element_recursive(driver, '//*[@id="upw"]')
                        pw_input.click(); pw_input.clear(); pw_input.send_keys(user_pw)
                        pw_input.send_keys(Keys.RETURN)
                        
                        self.log("⏳ 로그인 대기...")
                        # '바로가기' 버튼이 나타날 때까지 스마트 대기
                        # (로그인 성공 시 바로 넘어감)
                        for _ in range(20): # 0.5초 * 20 = 10초
                            # 팝업 닫기 (발견 즉시)
                            try:
                                driver.switch_to.default_content()
                                popups = driver.find_elements(By.XPATH, '//*[contains(@id, "closeBtn")]')
                                for btn in popups: 
                                    if btn.is_displayed(): driver.execute_script("arguments[0].click();", btn)
                            except: pass

                            # 로그인 성공 체크
                            if self.find_element_recursive(driver, XPATHS['MAIN_SHORTCUT']):
                                self.log("✅ 로그인 완료")
                                break
                            time.sleep(0.5)
                    else:
                        self.log("⚠️ 로그인 창 요소를 못 찾았습니다 (수동 로그인 필요)")

                except Exception as e: self.log(f"로그인 단계 패스: {e}")

            # ==========================================
            # B. 메뉴 이동 (스마트 대기 적용)
            # ==========================================
            if not target_window:
                self.log("📂 메뉴 이동...")
                try:
                    # (1) 바로가기 클릭
                    if wait_click(driver, XPATHS['MAIN_SHORTCUT']):
                        
                        # (2) HRM 클릭
                        # 메뉴 펼쳐짐 대기 (애니메이션 고려 짧게)
                        time.sleep(0.5) 
                        
                        # HRM 버튼 찾기 (반복 없이 스마트하게)
                        hrm_xpaths = ['//*[@id="BusinessSystem"]/li[2]', '//li[contains(text(), "HRM")]']
                        hrm_btn = None
                        for xpath in hrm_xpaths:
                            hrm_btn = self.find_element_recursive(driver, xpath)
                            if hrm_btn: break
                        
                        if hrm_btn:
                            cur_handles = driver.window_handles # 현재 창 개수 기억
                            driver.execute_script("arguments[0].click();", hrm_btn)
                            self.log("   👉 HRM 클릭")
                            
                            # (3) [핵심] 새 창이 뜰 때까지 대기 (스마트 웨이트)
                            # 창 개수가 늘어나면 즉시 통과! (sleep 제거)
                            WebDriverWait(driver, 10).until(EC.new_window_is_opened(cur_handles))
                            self.log("   ⚡ 새 창 열림 감지!")
                            
                            # 새 창으로 전환
                            new_handles = driver.window_handles
                            for h in new_handles:
                                if h not in cur_handles:
                                    driver.switch_to.window(h)
                                    break
                            
                            # (4) 근태 -> 특근신청 (요소 뜨자마자 클릭)
                            wait_click(driver, '//*[@id="gnrTopMenu_1_btnTopMenu"]/a')
                            wait_click(driver, '//*[@id="trvLeftMenu_label_4"]')
                            self.log("🎯 페이지 진입")
                            target_window = driver.current_window_handle
                        else:
                            self.log("❌ HRM 버튼 못 찾음")
                    else:
                        # 바로가기 버튼 못 찾음 -> 이미 특근창이 있을 수도? 재검사
                        pass

                except Exception as e:
                    self.log(f"⚠️ 메뉴 이동 중: {e}")

            # ==========================================
            # C. 작업 시작
            # ==========================================
            if not target_window:
                # 마지막으로 한번 더 찾기
                for handle in driver.window_handles:
                    driver.switch_to.window(handle)
                    driver.switch_to.default_content()
                    if self.find_element_recursive(driver, XPATHS["CALENDAR_ICON"]):
                        target_window = handle
                        break
            
            if not target_window:
                self.log("❌ 작업 대상 창을 못 찾았습니다.")
                self.btn_run.config(state='normal')
                return

            driver.switch_to.window(target_window)
            driver.switch_to.default_content()
            
            # --- 반복 입력 로직 (동일) ---
            target_dates = sorted(list(self.cal.selected_dates))
            success_cnt = 0
            dup_cnt = 0
            details = []

            for idx, date_str in enumerate(target_dates):
                self.log(f"▶ [{date_str}] 입력...")
                day_num = str(int(date_str.split('-')[2]))

                try:
                    # 달력 클릭 (뜨자마자)
                    if not wait_click(driver, XPATHS["CALENDAR_ICON"]):
                        self.log(f"⚠️ [{date_str}] 달력 아이콘 실패")
                        details.append(f"X {date_str}")
                        continue
                    
                    # 날짜 클릭 (달력 레이어 뜨는거 기다림 - 최대 2초)
                    found = False
                    try:
                        # 숫자 텍스트가 '보일 때'까지 기다림 -> 보이면 바로 클릭
                        # (XPath: 해당 텍스트를 가진 요소가 visible 상태가 되면)
                        xpath_day = f"//*[text()='{day_num}']"
                        # WebDriverWait은 재귀 검색이 안되므로, 기존 find_recursive를 짧은 주기로 호출
                        for _ in range(10): # 0.2초 * 10 = 2초
                            el = self.find_element_recursive(driver, xpath_day)
                            if el and el.is_displayed():
                                driver.execute_script("arguments[0].click();", el)
                                found = True
                                break
                            time.sleep(0.2)
                    except: pass
                    
                    if not found:
                        self.log(f"⚠️ [{date_str}] 날짜({day_num}) 못 찾음")
                        details.append(f"X {date_str}")
                        continue

                    # 시간/사유 입력 (첫번째만 or 매번)
                    if idx == 0:
                        wait_click(driver, XPATHS["START_H_BTN"])
                        wait_click(driver, f"//*[@id='cmbSTRT_HHXX_itemTable_{int(self.start_h_var.get())}']")
                        
                        # 입력칸은 클릭 대신 값 주입 (빠름)
                        el = self.find_element_recursive(driver, XPATHS["START_M_IPT"])
                        if el: 
                            driver.execute_script(f"arguments[0].value = '{self.start_m_var.get()}';", el)
                            driver.execute_script("arguments[0].dispatchEvent(new Event('change'));", el)

                        wait_click(driver, XPATHS["END_H_BTN"])
                        wait_click(driver, f"//*[@id='cmbENDX_HHXX_itemTable_{int(self.end_h_var.get())}']")
                        
                        el = self.find_element_recursive(driver, XPATHS["END_M_IPT"])
                        if el:
                            driver.execute_script(f"arguments[0].value = '{self.end_m_var.get()}';", el)
                            driver.execute_script("arguments[0].dispatchEvent(new Event('change'));", el)

                    # 사유 입력
                    el = self.find_element_recursive(driver, XPATHS["REASON_INPUT"])
                    if el: 
                        el.clear()
                        el.send_keys(self.reason_var.get())
                    
                    # 상신 클릭
                    wait_click(driver, XPATHS["SUBMIT_BTN"])
                    
                    # 결과 확인 (팝업이나 알럿이 뜨기를 스마트하게 기다림)
                    result_msg = "응답없음"
                    try:
                        # 0.1초 간격으로 알럿이나 메시지 뜸을 감지
                        start_t = time.time()
                        while time.time() - start_t < 3: # 최대 3초 대기
                            # 1. Alert 확인
                            try:
                                alert = driver.switch_to.alert
                                result_msg = alert.text
                                alert.accept()
                                break
                            except: pass
                            
                            # 2. HTML 메시지 확인
                            try:
                                msg_el = self.find_element_recursive(driver, XPATHS["DUPLICATE_MSG"])
                                if msg_el and msg_el.is_displayed():
                                    result_msg = msg_el.text.strip()
                                    wait_click(driver, XPATHS["POPUP_CONFIRM"])
                                    break
                            except: pass
                            time.sleep(0.1)
                    except: pass

                    # 로그 정리
                    if "정상" in result_msg or "완료" in result_msg:
                        self.log(f"✅ 성공")
                        success_cnt += 1
                        details.append(f"O {date_str}")
                    elif "중복" in result_msg or "이미" in result_msg:
                        self.log(f"⚠️ 중복")
                        dup_cnt += 1
                        details.append(f"X {date_str} (중복)")
                    else:
                        self.log(f"⛔ {result_msg}")
                        details.append(f"X {date_str} ({result_msg})")

                except Exception as e:
                    self.log(f"❌ 오류: {e}")
                    details.append(f"X {date_str} (에러)")

            # 최종 보고
            summary = f"성공 {success_cnt}, 중복 {dup_cnt}, 실패 {len(details)-success_cnt-dup_cnt}"
            self.log("🎉 " + summary)
            if details: messagebox.showinfo("완료", summary + "\n\n" + "\n".join(details))
            else: messagebox.showinfo("완료", "처리 내역 없음")

        except Exception as e:
            self.log(f"❌ 치명적 오류: {e}")
        finally:
            self.btn_run.config(state='normal')
if __name__ == "__main__":
    root = tk.Tk()
    app = AutoWorkApp(root)
    root.mainloop()