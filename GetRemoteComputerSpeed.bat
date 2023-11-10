@echo off
set REMOTE_COMPUTER=远程计算机的主机名或IP地址
set REMOTE_USER=远程计算机的用户名
set REMOTE_PASSWORD=远程计算机的密码
setlocal enabledelayedexpansion

for /f "tokens=1,* delims==" %%A in ('.\PsExec.exe \\%REMOTE_COMPUTER% -u %REMOTE_USER% -p %REMOTE_PASSWORD% cmd /c wmic NIC where "NetEnabled=true" get Name^,Speed /format:list') do (
    if not "%%B"=="" (
        if "%%A"=="Name" (
            set "name=%%B"
        ) else if "%%A"=="Speed" (
            set "speed=%%B"
            if not "!name!"=="" (
                REM 调用VBScript脚本进行转换
                for /f "delims=" %%S in ('cscript //nologo convertSpeed.vbs "!speed!"') do set "convertedSpeed=%%S"
                echo Name: !name!
                echo Speed: !convertedSpeed!
                echo.
                set "name="
            )
        )
    )
)

endlocal
pause