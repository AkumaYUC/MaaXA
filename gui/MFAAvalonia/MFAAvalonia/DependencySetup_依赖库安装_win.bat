@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul

set "ErrorOccurred=0"
set "UseWinget=1"

:: 检测系统架构 / Detect system architecture
set "ARCH=%PROCESSOR_ARCHITECTURE%"
if /i "%ARCH%"=="AMD64" (
    set "ARCH_NAME=x64"
    set "VC_PACKAGE=Microsoft.VCRedist.2015+.x64"
    set "VC_DOWNLOAD=https://aka.ms/vs/17/release/vc_redist.x64.exe"
    set "DOTNET_DOWNLOAD=https://aka.ms/dotnet/10.0/windowsdesktop-runtime-win-x64.exe"
) else if /i "%ARCH%"=="x86" (
    set "ARCH_NAME=x86"
    set "VC_PACKAGE=Microsoft.VCRedist.2015+.x86"
    set "VC_DOWNLOAD=https://aka.ms/vs/17/release/vc_redist.x86.exe"
    set "DOTNET_DOWNLOAD=https://aka.ms/dotnet/10.0/windowsdesktop-runtime-win-x86.exe"
) else if /i "%ARCH%"=="ARM64" (
    set "ARCH_NAME=ARM64"
    set "VC_PACKAGE=Microsoft.VCRedist.2015+.arm64"
    set "VC_DOWNLOAD=https://aka.ms/vs/17/release/vc_redist.arm64.exe"
    set "DOTNET_DOWNLOAD=https://aka.ms/dotnet/10.0/windowsdesktop-runtime-win-arm64.exe"
) else (
    echo [错误] 不支持的系统架构 / Unsupported architecture: %ARCH%
    echo.
    pause
    exit /b 1
)

echo.
echo 检测到系统架构 / Detected architecture: %ARCH_NAME%
echo.

:: 检测管理员权限 / Check administrator privileges
net session >nul 2>&1
if "%errorlevel%" neq "0" (
    echo ============================================================
    echo  需要管理员权限 / Administrator privileges required
    echo  请右键点击此脚本，选择"以管理员身份运行"
    echo  Right-click this file and choose "Run as administrator"
    echo ============================================================
    echo.
    pause
    exit /b 1
)

:: 检测 winget / Detect winget
where winget >nul 2>&1
if "%errorlevel%" neq "0" (
    set "UseWinget=0"
    echo winget 不可用，将使用直接下载方式 / winget unavailable, using direct download
    echo.
) else (
    winget --version >nul 2>&1
    if "!errorlevel!" neq "0" (
        set "UseWinget=0"
        echo winget 无法正常工作，将使用直接下载方式 / winget not working, using direct download
        echo.
    ) else (
        echo 检测到 winget 可用 / winget is available
        echo.
    )
)

set "TEMP_DIR=%TEMP%\MAADependencySetup"
if not exist "%TEMP_DIR%" mkdir "%TEMP_DIR%"

echo ============================================================
echo  正在安装 Microsoft Visual C++ Redistributable (%ARCH_NAME%)
echo  Installing VC++ Redistributable (%ARCH_NAME%)
echo ============================================================
echo.

if "%UseWinget%"=="1" (
    echo 首次使用 winget 可能提示接受协议，请输入 Y 并回车
    echo If prompted to accept agreements, type Y and press Enter.
    echo.
    winget install "%VC_PACKAGE%" --override "/install /passive /norestart" --uninstall-previous --accept-package-agreements --force
    if "!errorlevel!" neq "0" (
        set "ErrorOccurred=1"
        echo winget 安装失败，改用直接下载 / winget failed, trying direct download
        call :DownloadAndInstallVC
    )
) else (
    call :DownloadAndInstallVC
)

echo.
echo ============================================================
echo  正在安装 .NET Desktop Runtime 10.0 (%ARCH_NAME%)
echo  Installing .NET Desktop Runtime 10.0 (%ARCH_NAME%)
echo ============================================================
echo.

if "%UseWinget%"=="1" (
    winget install "Microsoft.DotNet.DesktopRuntime.10" --override "/install /passive /norestart" --uninstall-previous --accept-package-agreements --force
    if "!errorlevel!" neq "0" (
        set "ErrorOccurred=1"
        echo winget 安装失败，改用直接下载 / winget failed, trying direct download
        call :DownloadAndInstallDotNet
    )
) else (
    call :DownloadAndInstallDotNet
)

if exist "%TEMP_DIR%" rd /s /q "%TEMP_DIR%" >nul 2>&1

echo.
echo ============================================================
if "%ErrorOccurred%"=="0" (
    echo  运行库安装完成，请重启电脑后再次运行 MAA。
    echo  Installation complete. Please restart your PC and run MAA again.
) else (
    echo  安装过程中可能出现错误 / Errors may have occurred
    echo.
    echo  如失败，可手动复制下面链接到浏览器下载：
    echo  VC++ : %VC_DOWNLOAD%
    echo  .NET : %DOTNET_DOWNLOAD%
)
echo ============================================================
echo.
pause
exit /b 0

:DownloadAndInstallVC
set "VC_FILE=%TEMP_DIR%\vc_redist.exe"
echo 正在下载 VC++ Redistributable... / Downloading VC++...
certutil -urlcache -split -f "%VC_DOWNLOAD%" "%VC_FILE%" >nul 2>&1
if "!errorlevel!" neq "0" (
    echo certutil 下载失败，尝试 bitsadmin... / certutil failed, trying bitsadmin...
    bitsadmin /transfer "VCRedist" /priority high "%VC_DOWNLOAD%" "%VC_FILE%" >nul 2>&1
)
if exist "%VC_FILE%" (
    echo 正在安装 VC++ Redistributable... / Installing VC++...
    "%VC_FILE%" /install /passive /norestart
    set "RC=!errorlevel!"
    if not "!RC!"=="0" if not "!RC!"=="3010" (
        echo 安装失败 / Installation failed
        set "ErrorOccurred=1"
    ) else (
        echo 安装完成 / Installed
    )
    del "%VC_FILE%" >nul 2>&1
) else (
    echo 下载失败 / Download failed
    set "ErrorOccurred=1"
)
goto :eof

:DownloadAndInstallDotNet
set "DOTNET_FILE=%TEMP_DIR%\dotnet_runtime.exe"
echo 正在下载 .NET Desktop Runtime... / Downloading .NET...
certutil -urlcache -split -f "%DOTNET_DOWNLOAD%" "%DOTNET_FILE%" >nul 2>&1
if "!errorlevel!" neq "0" (
    echo certutil 下载失败，尝试 bitsadmin... / certutil failed, trying bitsadmin...
    bitsadmin /transfer "DotNetRuntime" /priority high "%DOTNET_DOWNLOAD%" "%DOTNET_FILE%" >nul 2>&1
)
if exist "%DOTNET_FILE%" (
    echo 正在安装 .NET Desktop Runtime... / Installing .NET...
    "%DOTNET_FILE%" /install /passive /norestart
    set "RC=!errorlevel!"
    if not "!RC!"=="0" if not "!RC!"=="3010" (
        echo 安装失败 / Installation failed
        set "ErrorOccurred=1"
    ) else (
        echo 安装完成 / Installed
    )
    del "%DOTNET_FILE%" >nul 2>&1
) else (
    echo 下载失败 / Download failed
    set "ErrorOccurred=1"
)
goto :eof
