@echo off
setlocal enabledelayedexpansion

echo ���ڰ�װ�Ҽ��˵�ѡ��...

:: ��ȡ��ǰ�ű�����Ŀ¼
set "currentPath=%~dp0"
set "toXlsbScriptPath=%currentPath%vbs_script\xlsxToxlsb.vbs"
set "toXlsxScriptPath=%currentPath%vbs_script\xlsbToxlsx.vbs"

:: ���VBS�ű��ļ��Ƿ����
if not exist "%toXlsbScriptPath%" (
    echo �����Ҳ���VBS�ű��ļ���%toXlsbScriptPath%
    pause
    exit /b 1
)

if not exist "%toXlsxScriptPath%" (
    echo �����Ҳ���VBS�ű��ļ���%toXlsxScriptPath%
    pause
    exit /b 1
)

:: ���xlsxתxlsbѡ��
reg add "HKEY_CLASSES_ROOT\SystemFileAssociations\.xlsx\shell\to_xlsb" /ve /d "to_xlsb" /f
reg add "HKEY_CLASSES_ROOT\SystemFileAssociations\.xlsx\shell\to_xlsb\command" /ve /d "wscript.exe \"%toXlsbScriptPath%\" \"%%1\"" /f
if %ERRORLEVEL% equ 0 (
    echo �ɹ����xlsxתxlsb�Ҽ��˵�ѡ��
) else (
    echo ���xlsxתxlsb�Ҽ��˵�ѡ��ʱ����
)

:: ���xlsbתxlsxѡ��
reg add "HKEY_CLASSES_ROOT\SystemFileAssociations\.xlsb\shell\to_xlsx" /ve /d "to_xlsx" /f
reg add "HKEY_CLASSES_ROOT\SystemFileAssociations\.xlsb\shell\to_xlsx\command" /ve /d "wscript.exe \"%toXlsxScriptPath%\" \"%%1\"" /f
if %ERRORLEVEL% equ 0 (
    echo �ɹ����xlsbתxlsx�Ҽ��˵�ѡ��
) else (
    echo ���xlsbתxlsx�Ҽ��˵�ѡ��ʱ����
)

echo ��װ��ɣ�
pause
