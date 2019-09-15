@echo off
::=========================  get administrator  =============================
title GET ADMIN
setlocal
set uac=~uac_permission_tmp_%random%
md "%SystemRoot%\system32\%uac%" 2>nul
if %errorlevel%==0 ( rd "%SystemRoot%\system32\%uac%" >nul 2>nul ) else (
    echo set uac = CreateObject^("Shell.Application"^)>"%temp%\%uac%.vbs"
    echo uac.ShellExecute "%~s0","","","runas",1 >>"%temp%\%uac%.vbs"
    echo WScript.Quit >>"%temp%\%uac%.vbs"
    "%temp%\%uac%.vbs" /f
    del /f /q "%temp%\%uac%.vbs" & exit
)
endlocal
::=========================  get administrator  =============================

::======================  env variable settings  ============================
set WinPath=%SystemRoot%\system32
set OfficeOPath=
set ProgramFilesPath=
wmic OS get OSArchitecture | find "32" >nul
if %errorlevel% equ 0 ( set ProgramFilesPath=%ProgramFiles^(x86^)%) else ( set ProgramFilesPath=%ProgramFiles%)
::======================  env variable settings  ============================

::===================  NCNU WINDOWS AND OFFICE ACT  =========================
title NCNU WINDOWS AND OFFICE ACT
mode con: cols=82 lines=40
:: "[介面] 檢查暨大校園網路環境 DNS Server Adderss: 163.22.2.1, 163.22.2.2" > 此取163.22.2.1
:ncnuenvcheck
cls
echo;
echo ==============================================================================
echo 　　　　　　　　         國立暨南大學授權軟體啟用程式
echo ==============================================================================
echo ●您好，正在檢查您是否在校園網路環境內...
(ipconfig /all | findstr "163.22.2.1") >nul || (ipconfig /all | findstr "ncnu.edu.tw") >nul
if %errorlevel% equ 0 (
    timeout /t 3
    goto actoptwhich
) else (
    timeout /t 3
    echo ●您好，您可能不在校園網路環境內，是否關閉程式？
    echo ^(1.^) 否，已重新更換網路環境，請重新檢查網路環境
    echo ^(2.^) 否，直接繼續使用本授權啟用程式
    echo ^(3.^) 是，系統將關閉程式，感謝您的使用
    choice /n /c 123
    if errorlevel 3 exit
    if errorlevel 2 goto actoptwhich
    if errorlevel 1 goto ncnuenvcheck
)

:: "[介面] 選擇 MICROSOFT WINDOWS 或 MICROSOFT OFFICE 進行校園授權認證"
:actoptwhich
cls
echo;
echo ==============================================================================
echo 　　　　　　　　         國立暨南大學授權軟體啟用程式
echo ==============================================================================
echo ●您好，請問您要選擇 Mircosoft Windows 或 Microsoft Office 進行啟用？
echo   ^(1.^) 啟用，Mircosoft Windows^(作業系統^)
echo   ^(2.^) 啟用，Microsoft Office^(辦公軟體套裝^)
echo   ^(3.^) 不啟用，系統將關閉程式，感謝您的使用
choice /n /c 123
if errorlevel 3 exit
if errorlevel 2 goto actoptoffice
if errorlevel 1 goto actoptwin

:: "[介面] 選擇 MICROSOFT WINDOWS 版本進行驗證"
:actoptwin
cls
echo;
echo ==============================================================================
echo 　　　　　　　      國立暨南大學 Mircosoft Windows 授權啟用驗證
echo ==============================================================================
echo ●您好，請問您要選擇 Mircosoft Windows 程式自動化版本驗證 或 手動選擇版本驗證 ？
echo ^(1.^) Mircosoft Windows 自動化版本驗證
echo ^(2.^) Mircosoft Windows 手動選擇版本驗證
echo ^(3.^) 返回選單，重新選擇 Mircosoft Windows 或 Microsoft Office 進行啟用
choice /n /c 123
if errorlevel 3 goto actoptwhich
if errorlevel 2 goto actwinmanual
if errorlevel 1 goto actwinauto

:actwinauto
cls
echo;
echo ==============================================================================
echo 　　　　　　　      國立暨南大學 Mircosoft Windows 授權啟用驗證
echo ==============================================================================
wmic os get Caption | find "Microsoft Windows 7" >nul
if %errorlevel% equ 0 ( goto actwin7set)
wmic os get Caption | find "Microsoft Windows 8" >nul
if %errorlevel% equ 0 ( goto actwin8set)
wmic os get Caption | find "Microsoft Windows 10" >nul
if %errorlevel% equ 0 ( goto actwin10set)
wmic os get Caption | find "Microsoft Windows Server 2012" >nul
if %errorlevel% equ 0 ( goto actwin2012set)
goto actoptwhich

:actwinmanual
cls
echo;
echo ==============================================================================
echo 　　　　　　　      國立暨南大學 Mircosoft Windows 授權啟用驗證
echo ==============================================================================
echo ●您好，請問您要選擇 Mircosoft Windows 哪一版本驗證？
echo ^(1.^) Mircosoft Windows 7
echo ^(2.^) Mircosoft Windows 8
echo ^(3.^) Mircosoft Windows 10
echo ^(4.^) Mircosoft Windows Server 2012
choice /n /c 1234
if errorlevel 4 ( goto actwin2012set )
if errorlevel 3 ( goto actwin10set )
if errorlevel 2 ( goto actwin8set )
if errorlevel 1 ( goto actwin7set )
goto actoptwhich

:actwin7set
wmic os get Caption | find "Microsoft Windows 7" >nul
if %errorlevel% equ 1 ( goto noinstallwin ) else (
    echo ●正在指定KMS主機中...
    cd %WinPath% & cscript slmgr.vbs /skms 10.10.6.5:1688
    goto actwinenable
)

:actwin8set
wmic os get Caption | find "Microsoft Windows 8" >nul
if %errorlevel% equ 1 ( goto noinstallwin ) else (
    echo ●正在指定KMS主機中...
    cd %WinPath% & cscript slmgr.vbs /skms 10.10.5.135:1688
    goto actwinenable
)

:actwin10set
wmic os get Caption | find "Microsoft Windows 10" >nul
if %errorlevel% equ 1 ( goto noinstallwin ) else (
    echo ●正在指定KMS主機中...
    cd %WinPath% & cscript slmgr.vbs /skms 10.10.5.135:1688
    goto actwinenable
)

:actwin2012set
wmic os get Caption | find "Microsoft Windows Server 2012" >nul
if %errorlevel% equ 1 ( goto noinstallwin ) else (
    echo ●正在指定KMS主機中...
    cd %WinPath% & cscript slmgr.vbs /skms 10.6.6.6:1688
    goto actwinenable
)

:actwinenable
echo ●啟用中...
cd %WinPath% & cscript slmgr.vbs /ato
if errorlevel 1 goto noinstallwin
echo ●自動查詢啟用狀態中...
cd %WinPath% & cscript slmgr.vbs /dlv
if errorlevel 1 goto noinstallwin
echo        ● [Mircosoft Windows] 啟用完成，請按鍵盤任意鍵繼續 ●
echo ==============================================================================
pause > nul
goto actoptwhich

:: "[介面] 選擇 MICROSOFT OFFICE 版本進行驗證"
:actoptoffice
cls
echo;
echo ==============================================================================
echo 　　　　　　　      國立暨南大學 Mircosoft Office 授權啟用驗證
echo ==============================================================================
echo ●您好，請問您要選擇 Microsoft Office 程式自動化版本驗證 或 手動選擇版本驗證 ？
echo ^(1.^) Microsoft Office 自動化版本驗證
echo ^(2.^) Microsoft Office 手動選擇版本驗證
echo ^(3.^) 返回選單，重新選擇 Mircosoft Windows 或 Microsoft Office 進行啟用
choice /n /c 123
if errorlevel 3 goto actoptwhich
if errorlevel 2 goto actofficemanual
if errorlevel 1 goto actofficeauto
goto actoptwhich

:actofficeauto
cls
echo;
echo ==============================================================================
echo 　　　　　　　      國立暨南大學 Mircosoft Office 授權啟用驗證
echo ==============================================================================
if exist "%ProgramFilesPath%\Microsoft Office\Office13" ( goto actoffice2010 )
if exist "%ProgramFilesPath%\Microsoft Office\Office14" ( goto actoffice2013 )
if exist "%ProgramFilesPath%\Microsoft Office\Office15" ( goto actoffice2016 )
if exist "%ProgramFilesPath%\Microsoft Office\Office16" ( goto actoffice2019 )
goto actoptwhich

:actofficemanual
cls
echo;
echo ==============================================================================
echo 　　　　　　　      國立暨南大學 Mircosoft Office 授權啟用驗證
echo ==============================================================================
echo ●您好，請問您要選擇 Mircosoft Office 哪一版本驗證？
echo ^(1.^) Mircosoft Windows Office 2010
echo ^(2.^) Mircosoft Windows Office 2013
echo ^(3.^) Mircosoft Windows Office 2016
echo ^(4.^) Mircosoft Windows Office 2019
choice /n /c 1234
if errorlevel 4 ( goto actoffice2019 )
if errorlevel 3 ( goto actoffice2016 )
if errorlevel 2 ( goto actoffice2013 )
if errorlevel 1 ( goto actoffice2010 )
goto actoptwhich

:actoffice2010
if not exist "%ProgramFilesPath%\Microsoft Office\Office13" ( goto noinstalloffice )
echo ●正在指定KMS主機中...
set OfficeOPath=%ProgramFilesPath%\Microsoft Office\Office13
cd %WinPath% & cscript "%OfficeOPath%\ospp.vbs" /sethst:10.6.6.6
cd %WinPath% & cscript "%OfficeOPath%\ospp.vbs" /setprt:1688
goto actofficeenable

:actoffice2013
if not exist "%ProgramFilesPath%\Microsoft Office\Office14" ( goto noinstalloffice )
echo ●正在指定KMS主機中...
set OfficeOPath=%ProgramFilesPath%\Microsoft Office\Office14
cd %WinPath% & cscript "%OfficeOPath%\ospp.vbs" /sethst:10.10.5.135
cd %WinPath% & cscript "%OfficeOPath%\ospp.vbs" /setprt:1688
goto actofficeenable

:actoffice2016
if not exist "%ProgramFilesPath%\Microsoft Office\Office15" ( goto noinstalloffice )
echo ●正在指定KMS主機中...
set OfficeOPath=%ProgramFilesPath%\Microsoft Office\Office15
cd %WinPath% & cscript "%OfficeOPath%\ospp.vbs" /sethst:10.10.5.135
cd %WinPath% & cscript "%OfficeOPath%\ospp.vbs" /setprt:1688
goto actofficeenable

:actoffice2019
if not exist "%ProgramFilesPath%\Microsoft Office\Office16" ( goto noinstalloffice )
echo ●正在指定KMS主機中...
set OfficeOPath=%ProgramFilesPath%\Microsoft Office\Office16
cd %WinPath% & cscript "%OfficeOPath%\ospp.vbs" /sethst:10.10.5.135
cd %WinPath% & cscript "%OfficeOPath%\ospp.vbs" /setprt:1688
goto actofficeenable

:actofficeenable
echo ●啟用中...
cd %WinPath% & cscript "%OfficeOPath%\ospp.vbs" /act
REM cd %WinPath% & cscript "%OfficeOPath%\ospp.vbs" /osppsvcrestart
if errorlevel 1 goto noinstalloffice
echo ●自動查詢啟用狀態中...
cd %WinPath% & cscript "%OfficeOPath%\ospp.vbs" /dstatus
if errorlevel 1 goto noinstalloffice
echo        ● [Mircosoft Office] 啟用完成，請按鍵盤任意鍵繼續 ●
echo ==============================================================================
pause > nul
goto actoptwhich

:noinstallwin
REM cls
echo;
echo ==============================================================================
echo 　　　　　　　      國立暨南大學 Mircosoft Windows 授權啟用驗證^(錯誤^)
echo ==============================================================================
echo ●您好，是否版本選擇錯誤？若出現「輸入錯誤：無法找到script檔」
echo  ^(請建議您選用自動驗證模式，或進行手動，遵循操作說明啟用驗證^)
echo;
echo ●您好，您是不是跳出錯誤視窗了呢？
echo  ^(請您至 [暨大全校授權軟體下載系統]→[檢視軟體說明及序號]→[選擇相對應版本]，進行手動啟用^)
echo;
echo ●您好，您是否於啟用驗證過程中，出現錯誤代碼"ERROR CODE"？
echo  ^(請您於稍後開啟的[常見問題頁面]，參考錯誤代碼"ERROR CODE"，進行相關問題排除^)
echo;
echo ==============================================================================
echo        ●請按鍵盤任意鍵開啟 [常見問題頁面] 和 [暨大全校授權軟體下載系統] ●
echo ==============================================================================
pause >nul
start iexplore.exe "http://wiki.kmu.edu.tw/index.php/%%E5%%B8%%B8%%E8%%A6%%8B_KMS_%%E5%%95%%9F%%E7%%94%%A8%%E5%%95%%8F%%E9%%A1%%8C"
start iexplore.exe "http://ccweb.ncnu.edu.tw/softlib/software_download_date_statlist.cshtml"
goto actoptwhich

:noinstalloffice
REM cls
echo;
echo ==============================================================================
echo 　　　　　　　      國立暨南大學 Mircosoft Office 授權啟用驗證^(錯誤^)
echo ==============================================================================
echo ●您好，是否版本選擇錯誤？若出現「輸入錯誤：無法找到script檔」
echo  ^(請建議您選用自動驗證模式，或進行手動，遵循操作說明啟用驗證^)
echo;
echo ●您好，您是否尚未安裝 Microsoft Office^(辦公軟體套裝^) ？
echo  ^(請您至 [控制台]→[程式集]→[程式和功能] 檢查安裝狀況^)
echo;
echo ●您好，您是否於啟用驗證過程中，出現錯誤代碼"ERROR CODE"？
echo  ^(請您於稍後開啟的[常見問題頁面]，參考錯誤代碼"ERROR CODE"，進行相關問題排除^)
echo;
echo ●您好，您是否已安裝 Microsoft Office，但安裝初始並非安裝在預設路徑中？
echo  ^(請您至 [暨大全校授權軟體下載系統]，參考檢視軟體說明及序號，進行手動啟用^)
echo;
echo ●您好，您若無法正常開啟[常見問題頁面]，請先行檢查該台電腦是否網路不通
echo ==============================================================================
echo        ●請按鍵盤任意鍵開啟 [常見問題頁面] 和 [暨大全校授權軟體下載系統] ●
echo ==============================================================================
pause >nul
start iexplore.exe "http://wiki.kmu.edu.tw/index.php/%%E5%%B8%%B8%%E8%%A6%%8B_KMS_%%E5%%95%%9F%%E7%%94%%A8%%E5%%95%%8F%%E9%%A1%%8C"
start iexplore.exe "http://ccweb.ncnu.edu.tw/softlib/software_download_date_statlist.cshtml"
goto actoptwhich
::===================  NCNU WINDOWS AND OFFICE ACT  =========================
::systeminfo >"%userprofile%\desktop\systeminfo.log"
::wmic os get Caption, OSArchitecture