@echo off
rem ===================================================================
rem  FolderPad 安装 / 卸载 / 自检（双击即可）
rem    1 = 安装桌面快捷方式 + 发送到入口
rem    2 = 卸载快捷方式
rem    3 = 运行自检
rem    4 = 带日志启动（调试用，会保留控制台窗口）
rem ===================================================================
setlocal
cd /d "%~dp0"
title FolderPad

if "%~1"=="" goto menu
if /i "%~1"=="install" goto install
if /i "%~1"=="uninstall" goto uninstall
if /i "%~1"=="selftest" goto selftest
if /i "%~1"=="debug" goto debug
goto menu

:menu
cls
echo.
echo   FolderPad —— 桌面围栏式文件夹框
echo   ---------------------------------
echo   1. 安装（桌面快捷方式 + 右键「发送到」入口）
echo   2. 卸载快捷方式（不动数据、不动文件）
echo   3. 运行自检
echo   4. 调试启动（显示控制台与错误）
echo   5. 退出
echo.
set /p choice=请选择 [1-5]: 
if "%choice%"=="1" goto install
if "%choice%"=="2" goto uninstall
if "%choice%"=="3" goto selftest
if "%choice%"=="4" goto debug
goto :eof

:install
powershell.exe -NoProfile -Sta -ExecutionPolicy Bypass -File "%~dp0FolderPad.ps1" -Install
echo.
pause
goto :eof

:uninstall
powershell.exe -NoProfile -Sta -ExecutionPolicy Bypass -File "%~dp0FolderPad.ps1" -Uninstall
echo.
pause
goto :eof

:selftest
powershell.exe -NoProfile -Sta -ExecutionPolicy Bypass -File "%~dp0FolderPad.ps1" -SelfTest
echo.
pause
goto :eof

:debug
powershell.exe -NoProfile -Sta -ExecutionPolicy Bypass -File "%~dp0FolderPad.ps1" %2 %3 %4
echo.
echo 程序已退出。
pause
goto :eof
