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
:: "[����] �ˬd�[�j�ն�������� DNS Server Adderss: 163.22.2.1, 163.22.2.2" > ����163.22.2.1
:ncnuenvcheck
cls
echo;
echo ==============================================================================
echo �@�@�@�@�@�@�@�@         ��ߺ[�n�j�Ǳ��v�n��ҥε{��
echo ==============================================================================
echo ���z�n�A���b�ˬd�z�O�_�b�ն�������Ҥ�...
(ipconfig /all | findstr "163.22.2.1") >nul || (ipconfig /all | findstr "ncnu.edu.tw") >nul
if %errorlevel% equ 0 (
    timeout /t 3
    goto actoptwhich
) else (
    timeout /t 3
    echo ���z�n�A�z�i�ण�b�ն�������Ҥ��A�O�_�����{���H
    echo ^(1.^) �_�A�w���s�󴫺������ҡA�Э��s�ˬd��������
    echo ^(2.^) �_�A�����~��ϥΥ����v�ҥε{��
    echo ^(3.^) �O�A�t�αN�����{���A�P�±z���ϥ�
    choice /n /c 123
    if errorlevel 3 exit
    if errorlevel 2 goto actoptwhich
    if errorlevel 1 goto ncnuenvcheck
)

:: "[����] ��� MICROSOFT WINDOWS �� MICROSOFT OFFICE �i��ն���v�{��"
:actoptwhich
cls
echo;
echo ==============================================================================
echo �@�@�@�@�@�@�@�@         ��ߺ[�n�j�Ǳ��v�n��ҥε{��
echo ==============================================================================
echo ���z�n�A�аݱz�n��� Mircosoft Windows �� Microsoft Office �i��ҥΡH
echo   ^(1.^) �ҥΡAMircosoft Windows^(�@�~�t��^)
echo   ^(2.^) �ҥΡAMicrosoft Office^(�줽�n��M��^)
echo   ^(3.^) ���ҥΡA�t�αN�����{���A�P�±z���ϥ�
choice /n /c 123
if errorlevel 3 exit
if errorlevel 2 goto actoptoffice
if errorlevel 1 goto actoptwin

:: "[����] ��� MICROSOFT WINDOWS �����i������"
:actoptwin
cls
echo;
echo ==============================================================================
echo �@�@�@�@�@�@�@      ��ߺ[�n�j�� Mircosoft Windows ���v�ҥ�����
echo ==============================================================================
echo ���z�n�A�аݱz�n��� Mircosoft Windows �{���۰ʤƪ������� �� ��ʿ�ܪ������� �H
echo ^(1.^) Mircosoft Windows �۰ʤƪ�������
echo ^(2.^) Mircosoft Windows ��ʿ�ܪ�������
echo ^(3.^) ��^���A���s��� Mircosoft Windows �� Microsoft Office �i��ҥ�
choice /n /c 123
if errorlevel 3 goto actoptwhich
if errorlevel 2 goto actwinmanual
if errorlevel 1 goto actwinauto

:actwinauto
cls
echo;
echo ==============================================================================
echo �@�@�@�@�@�@�@      ��ߺ[�n�j�� Mircosoft Windows ���v�ҥ�����
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
echo �@�@�@�@�@�@�@      ��ߺ[�n�j�� Mircosoft Windows ���v�ҥ�����
echo ==============================================================================
echo ���z�n�A�аݱz�n��� Mircosoft Windows ���@�������ҡH
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
    echo �����b���wKMS�D����...
    cd %WinPath% & cscript slmgr.vbs /skms 10.10.6.5:1688
    goto actwinenable
)

:actwin8set
wmic os get Caption | find "Microsoft Windows 8" >nul
if %errorlevel% equ 1 ( goto noinstallwin ) else (
    echo �����b���wKMS�D����...
    cd %WinPath% & cscript slmgr.vbs /skms 10.10.5.135:1688
    goto actwinenable
)

:actwin10set
wmic os get Caption | find "Microsoft Windows 10" >nul
if %errorlevel% equ 1 ( goto noinstallwin ) else (
    echo �����b���wKMS�D����...
    cd %WinPath% & cscript slmgr.vbs /skms 10.10.5.135:1688
    goto actwinenable
)

:actwin2012set
wmic os get Caption | find "Microsoft Windows Server 2012" >nul
if %errorlevel% equ 1 ( goto noinstallwin ) else (
    echo �����b���wKMS�D����...
    cd %WinPath% & cscript slmgr.vbs /skms 10.6.6.6:1688
    goto actwinenable
)

:actwinenable
echo ���ҥΤ�...
cd %WinPath% & cscript slmgr.vbs /ato
if errorlevel 1 goto noinstallwin
echo ���۰ʬd�߱ҥΪ��A��...
cd %WinPath% & cscript slmgr.vbs /dlv
if errorlevel 1 goto noinstallwin
echo        �� [Mircosoft Windows] �ҥΧ����A�Ы���L���N���~�� ��
echo ==============================================================================
pause > nul
goto actoptwhich

:: "[����] ��� MICROSOFT OFFICE �����i������"
:actoptoffice
cls
echo;
echo ==============================================================================
echo �@�@�@�@�@�@�@      ��ߺ[�n�j�� Mircosoft Office ���v�ҥ�����
echo ==============================================================================
echo ���z�n�A�аݱz�n��� Microsoft Office �{���۰ʤƪ������� �� ��ʿ�ܪ������� �H
echo ^(1.^) Microsoft Office �۰ʤƪ�������
echo ^(2.^) Microsoft Office ��ʿ�ܪ�������
echo ^(3.^) ��^���A���s��� Mircosoft Windows �� Microsoft Office �i��ҥ�
choice /n /c 123
if errorlevel 3 goto actoptwhich
if errorlevel 2 goto actofficemanual
if errorlevel 1 goto actofficeauto
goto actoptwhich

:actofficeauto
cls
echo;
echo ==============================================================================
echo �@�@�@�@�@�@�@      ��ߺ[�n�j�� Mircosoft Office ���v�ҥ�����
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
echo �@�@�@�@�@�@�@      ��ߺ[�n�j�� Mircosoft Office ���v�ҥ�����
echo ==============================================================================
echo ���z�n�A�аݱz�n��� Mircosoft Office ���@�������ҡH
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
echo �����b���wKMS�D����...
set OfficeOPath=%ProgramFilesPath%\Microsoft Office\Office13
cd %WinPath% & cscript "%OfficeOPath%\ospp.vbs" /sethst:10.6.6.6
cd %WinPath% & cscript "%OfficeOPath%\ospp.vbs" /setprt:1688
goto actofficeenable

:actoffice2013
if not exist "%ProgramFilesPath%\Microsoft Office\Office14" ( goto noinstalloffice )
echo �����b���wKMS�D����...
set OfficeOPath=%ProgramFilesPath%\Microsoft Office\Office14
cd %WinPath% & cscript "%OfficeOPath%\ospp.vbs" /sethst:10.10.5.135
cd %WinPath% & cscript "%OfficeOPath%\ospp.vbs" /setprt:1688
goto actofficeenable

:actoffice2016
if not exist "%ProgramFilesPath%\Microsoft Office\Office15" ( goto noinstalloffice )
echo �����b���wKMS�D����...
set OfficeOPath=%ProgramFilesPath%\Microsoft Office\Office15
cd %WinPath% & cscript "%OfficeOPath%\ospp.vbs" /sethst:10.10.5.135
cd %WinPath% & cscript "%OfficeOPath%\ospp.vbs" /setprt:1688
goto actofficeenable

:actoffice2019
if not exist "%ProgramFilesPath%\Microsoft Office\Office16" ( goto noinstalloffice )
echo �����b���wKMS�D����...
set OfficeOPath=%ProgramFilesPath%\Microsoft Office\Office16
cd %WinPath% & cscript "%OfficeOPath%\ospp.vbs" /sethst:10.10.5.135
cd %WinPath% & cscript "%OfficeOPath%\ospp.vbs" /setprt:1688
goto actofficeenable

:actofficeenable
echo ���ҥΤ�...
cd %WinPath% & cscript "%OfficeOPath%\ospp.vbs" /act
REM cd %WinPath% & cscript "%OfficeOPath%\ospp.vbs" /osppsvcrestart
if errorlevel 1 goto noinstalloffice
echo ���۰ʬd�߱ҥΪ��A��...
cd %WinPath% & cscript "%OfficeOPath%\ospp.vbs" /dstatus
if errorlevel 1 goto noinstalloffice
echo        �� [Mircosoft Office] �ҥΧ����A�Ы���L���N���~�� ��
echo ==============================================================================
pause > nul
goto actoptwhich

:noinstallwin
REM cls
echo;
echo ==============================================================================
echo �@�@�@�@�@�@�@      ��ߺ[�n�j�� Mircosoft Windows ���v�ҥ�����^(���~^)
echo ==============================================================================
echo ���z�n�A�O�_������ܿ��~�H�Y�X�{�u��J���~�G�L�k���script�ɡv
echo  ^(�Ы�ĳ�z��Φ۰����ҼҦ��A�ζi���ʡA��`�ާ@�����ҥ�����^)
echo;
echo ���z�n�A�z�O���O���X���~�����F�O�H
echo  ^(�бz�� [�[�j���ձ��v�n��U���t��]��[�˵��n�黡���ΧǸ�]��[��ܬ۹�������]�A�i���ʱҥ�^)
echo;
echo ���z�n�A�z�O�_��ҥ����ҹL�{���A�X�{���~�N�X"ERROR CODE"�H
echo  ^(�бz��y��}�Ҫ�[�`�����D����]�A�Ѧҿ��~�N�X"ERROR CODE"�A�i��������D�ư�^)
echo;
echo ==============================================================================
echo        ���Ы���L���N��}�� [�`�����D����] �M [�[�j���ձ��v�n��U���t��] ��
echo ==============================================================================
pause >nul
start iexplore.exe "http://wiki.kmu.edu.tw/index.php/%%E5%%B8%%B8%%E8%%A6%%8B_KMS_%%E5%%95%%9F%%E7%%94%%A8%%E5%%95%%8F%%E9%%A1%%8C"
start iexplore.exe "http://ccweb.ncnu.edu.tw/softlib/software_download_date_statlist.cshtml"
goto actoptwhich

:noinstalloffice
REM cls
echo;
echo ==============================================================================
echo �@�@�@�@�@�@�@      ��ߺ[�n�j�� Mircosoft Office ���v�ҥ�����^(���~^)
echo ==============================================================================
echo ���z�n�A�O�_������ܿ��~�H�Y�X�{�u��J���~�G�L�k���script�ɡv
echo  ^(�Ы�ĳ�z��Φ۰����ҼҦ��A�ζi���ʡA��`�ާ@�����ҥ�����^)
echo;
echo ���z�n�A�z�O�_�|���w�� Microsoft Office^(�줽�n��M��^) �H
echo  ^(�бz�� [����x]��[�{����]��[�{���M�\��] �ˬd�w�˪��p^)
echo;
echo ���z�n�A�z�O�_��ҥ����ҹL�{���A�X�{���~�N�X"ERROR CODE"�H
echo  ^(�бz��y��}�Ҫ�[�`�����D����]�A�Ѧҿ��~�N�X"ERROR CODE"�A�i��������D�ư�^)
echo;
echo ���z�n�A�z�O�_�w�w�� Microsoft Office�A���w�˪�l�ëD�w�˦b�w�]���|���H
echo  ^(�бz�� [�[�j���ձ��v�n��U���t��]�A�Ѧ��˵��n�黡���ΧǸ��A�i���ʱҥ�^)
echo;
echo ���z�n�A�z�Y�L�k���`�}��[�`�����D����]�A�Х����ˬd�ӥx�q���O�_�������q
echo ==============================================================================
echo        ���Ы���L���N��}�� [�`�����D����] �M [�[�j���ձ��v�n��U���t��] ��
echo ==============================================================================
pause >nul
start iexplore.exe "http://wiki.kmu.edu.tw/index.php/%%E5%%B8%%B8%%E8%%A6%%8B_KMS_%%E5%%95%%9F%%E7%%94%%A8%%E5%%95%%8F%%E9%%A1%%8C"
start iexplore.exe "http://ccweb.ncnu.edu.tw/softlib/software_download_date_statlist.cshtml"
goto actoptwhich
::===================  NCNU WINDOWS AND OFFICE ACT  =========================
::systeminfo >"%userprofile%\desktop\systeminfo.log"
::wmic os get Caption, OSArchitecture