; =========================================================
; [최초 설치용] Special Work Writer Setup (이름 정리 완료)
; =========================================================

; 1. [라이브러리 불러오기]
!include "MUI2.nsh"

; 2. [프로그램 기본 정보]
Name "특근 입력기"
OutFile "특근입력기 설치 프로그램.exe"                ; 이 스크립트가 만들어낼 '설치 파일' 이름
InstallDir "$PROGRAMFILES\Special_Work_Writer" ; 설치될 폴더 경로
InstallDirRegKey HKLM "Software\Special_Work_Writer" "Install_Dir"
RequestExecutionLevel admin

; 3. [디자인/아이콘 설정]
!define MUI_ABORTWARNING
!define MUI_ICON "jjangu3.ico"
!define MUI_UNICON "jjangu3.ico"

; 4. [페이지 순서 설정]
!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_INSTFILES
!define MUI_FINISHPAGE_RUN "$INSTDIR\Special_Work_Writer.exe"
!define MUI_FINISHPAGE_RUN_TEXT "특근 입력기 바로 실행하기"
!insertmacro MUI_PAGE_FINISH

; 5. [언어 설정]
!insertmacro MUI_LANGUAGE "Korean"

; =========================================================
; [메인] 설치 섹션
; =========================================================
Section "MainSection" SecMain

    SetOutPath "$INSTDIR"

    ; 1. 기존 프로그램 종료 (이름 변경 반영: 8 제거)
    ExecWait "taskkill /F /IM Special_Work_Writer.exe"

    ; 2. 파일 복사
    ; ★중요★: 갖고 계신 exe 파일 이름을 'Special_Work_Writer.exe'로 변경해주세요!
    File "Special_Work_Writer.exe"

    ; 3. 레지스트리 등록
    WriteRegStr HKLM "Software\Special_Work_Writer" "Install_Dir" "$INSTDIR"

    ; 4. 프로그램 추가/제거 목록 등록
    WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Special_Work_Writer" "DisplayName" "특근 입력기"
    WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Special_Work_Writer" "UninstallString" '"$INSTDIR\uninstall.exe"'
    WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Special_Work_Writer" "DisplayIcon" '"$INSTDIR\Special_Work_Writer.exe"'

    ; 5. 언인스톨러 생성
    WriteUninstaller "$INSTDIR\uninstall.exe"

    ; 6. 바로가기 생성
    CreateShortCut "$DESKTOP\특근입력기.lnk" "$INSTDIR\Special_Work_Writer.exe"

    CreateDirectory "$SMPROGRAMS\특근입력기"
    CreateShortCut "$SMPROGRAMS\특근입력기\특근입력기.lnk" "$INSTDIR\Special_Work_Writer.exe"
    CreateShortCut "$SMPROGRAMS\특근입력기\제거하기.lnk" "$INSTDIR\uninstall.exe"

SectionEnd

; =========================================================
; [제거] 삭제 섹션
; =========================================================
Section "Uninstall"
    ; 프로세스 종료 (이름 변경 반영)
    ExecWait "taskkill /F /IM Special_Work_Writer.exe"

    ; 파일 삭제
    Delete "$INSTDIR\Special_Work_Writer.exe"
    Delete "$INSTDIR\uninstall.exe"

    ; 바로가기 삭제
    Delete "$DESKTOP\특근입력기.lnk"
    RMDir /r "$SMPROGRAMS\특근입력기"
    RMDir "$INSTDIR"

    ; 레지스트리 삭제
    DeleteRegKey HKLM "Software\Special_Work_Writer"
    DeleteRegKey HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Special_Work_Writer"
SectionEnd