@echo off
setlocal enabledelayedexpansion

for /f "tokens=1,* delims==" %%A in ('wmic NIC where "NetEnabled=true" get Name^,Speed /format:list') do (
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