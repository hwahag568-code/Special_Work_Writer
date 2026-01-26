; =========================================================
; [업데이트용] Special Work Writer Update
; =========================================================

!include "MUI2.nsh"

; 1. 프로그램 정보
Name "특근 입력기 업데이트"
OutFile "Update_Work_Writer.exe" ; ★이 이름으로 깃허브에 올라가야 함
RequestExecutionLevel admin

; 디자인 (아이콘만 설정)
!define MUI_ICON "jjangu3.ico"

; 2. 설치 페이지 (업데이트 때는 '설치 중' 게이지바만 보여줌)
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_LANGUAGE "Korean"

Section "Update"
    ; 1. 설치된 경로 찾아오기 (최초 설치 때 레지스트리에 저장한 값)
    ReadRegStr $0 HKLM "Software\Special_Work_Writer" "Install_Dir"

    ; 만약 레지스트리가 없으면(설치 안 된 PC면) 기본 경로로 강제 설정
    ${If} $0 == ""
        StrCpy $0 "$PROGRAMFILES\Special_Work_Writer"
    ${EndIf}

    ; 2. 실행 중인 내 프로그램 강제 종료 (가장 중요 ?)
    ; 크롬드라이버도 같이 죽입니다.
    ExecWait "taskkill /F /IM Special_Work_Writer.exe"
    ExecWait "taskkill /F /IM chromedriver.exe"

    ; 잠깐 대기 (파일 잠금 해제 시간 확보)
    Sleep 1000

    ; 3. 파일 덮어쓰기
    SetOutPath "$0"
    File "Special_Work_Writer.exe"
    ; 아이콘 등은 안 바뀌었으면 굳이 안 넣어도 되지만, 안전하게 포함

    ; 4. 업데이트 완료 후 프로그램 자동 재실행
    Exec "$0\Special_Work_Writer.exe"

SectionEnd