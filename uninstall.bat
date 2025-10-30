@echo off
setlocal enabledelayedexpansion

echo 正在卸载右键菜单选项...

:: 删除xlsx转xlsb选项
reg delete "HKEY_CLASSES_ROOT\SystemFileAssociations\.xlsx\shell\to_xlsb\command" /f >nul 2>&1
reg delete "HKEY_CLASSES_ROOT\SystemFileAssociations\.xlsx\shell\to_xlsb" /f >nul 2>&1
if %ERRORLEVEL% equ 0 (
    echo 成功移除xlsx转xlsb右键菜单选项
) else (
    echo 移除xlsx转xlsb右键菜单选项时出错或选项不存在
)

:: 删除xlsb转xlsx选项
reg delete "HKEY_CLASSES_ROOT\SystemFileAssociations\.xlsb\shell\to_xlsx\command" /f >nul 2>&1
reg delete "HKEY_CLASSES_ROOT\SystemFileAssociations\.xlsb\shell\to_xlsx" /f >nul 2>&1
if %ERRORLEVEL% equ 0 (
    echo 成功移除xlsb转xlsx右键菜单选项
) else (
    echo 移除xlsb转xlsx右键菜单选项时出错或选项不存在
)

echo 卸载完成！
pause
