@echo off
setlocal enabledelayedexpansion

echo ����ж���Ҽ��˵�ѡ��...

:: ɾ��xlsxתxlsbѡ��
reg delete "HKEY_CLASSES_ROOT\SystemFileAssociations\.xlsx\shell\to_xlsb\command" /f >nul 2>&1
reg delete "HKEY_CLASSES_ROOT\SystemFileAssociations\.xlsx\shell\to_xlsb" /f >nul 2>&1
if %ERRORLEVEL% equ 0 (
    echo �ɹ��Ƴ�xlsxתxlsb�Ҽ��˵�ѡ��
) else (
    echo �Ƴ�xlsxתxlsb�Ҽ��˵�ѡ��ʱ�����ѡ�����
)

:: ɾ��xlsbתxlsxѡ��
reg delete "HKEY_CLASSES_ROOT\SystemFileAssociations\.xlsb\shell\to_xlsx\command" /f >nul 2>&1
reg delete "HKEY_CLASSES_ROOT\SystemFileAssociations\.xlsb\shell\to_xlsx" /f >nul 2>&1
if %ERRORLEVEL% equ 0 (
    echo �ɹ��Ƴ�xlsbתxlsx�Ҽ��˵�ѡ��
) else (
    echo �Ƴ�xlsbתxlsx�Ҽ��˵�ѡ��ʱ�����ѡ�����
)

echo ж����ɣ�
pause
