@echo off
setlocal enabledelayedexpansion

echo 正在安装右键菜单选项...

:: 获取当前脚本所在目录
set "currentPath=%~dp0"
set "toXlsbScriptPath=%currentPath%vbs_script\xlsxToxlsb.vbs"
set "toXlsxScriptPath=%currentPath%vbs_script\xlsbToxlsx.vbs"

:: 检查VBS脚本文件是否存在
if not exist "%toXlsbScriptPath%" (
    echo 错误：找不到VBS脚本文件：%toXlsbScriptPath%
    pause
    exit /b 1
)

if not exist "%toXlsxScriptPath%" (
    echo 错误：找不到VBS脚本文件：%toXlsxScriptPath%
    pause
    exit /b 1
)

:: 添加xlsx转xlsb选项
reg add "HKEY_CLASSES_ROOT\SystemFileAssociations\.xlsx\shell\to_xlsb" /ve /d "to_xlsb" /f
reg add "HKEY_CLASSES_ROOT\SystemFileAssociations\.xlsx\shell\to_xlsb\command" /ve /d "wscript.exe \"%toXlsbScriptPath%\" \"%%1\"" /f
if %ERRORLEVEL% equ 0 (
    echo 成功添加xlsx转xlsb右键菜单选项
) else (
    echo 添加xlsx转xlsb右键菜单选项时出错
)

:: 添加xlsb转xlsx选项
reg add "HKEY_CLASSES_ROOT\SystemFileAssociations\.xlsb\shell\to_xlsx" /ve /d "to_xlsx" /f
reg add "HKEY_CLASSES_ROOT\SystemFileAssociations\.xlsb\shell\to_xlsx\command" /ve /d "wscript.exe \"%toXlsxScriptPath%\" \"%%1\"" /f
if %ERRORLEVEL% equ 0 (
    echo 成功添加xlsb转xlsx右键菜单选项
) else (
    echo 添加xlsb转xlsx右键菜单选项时出错
)

echo 安装完成！
pause
