@echo off
chcp 65001 >nul
echo ============================================================
echo   涌益现货数据 → 云端一键同步
echo ============================================================
echo.

:: 1. 扫描桌面最新 Excel
set "LATEST="
for /f "delims=" %%f in ('dir /b /o-d "D:\CC\Desktop\*涌益咨询日度数据*.xlsx" 2^>nul') do (
    set "LATEST=D:\CC\Desktop\%%f"
    goto :found
)
for /f "delims=" %%f in ('dir /b /o-d "D:\CC\Desktop\*涌益咨询*.xlsx" 2^>nul') do (
    set "LATEST=D:\CC\Desktop\%%f"
    goto :found
)
echo [错误] 桌面上没有找到涌益咨询 Excel 文件！
pause
exit /b 1

:found
echo [1/3] 找到最新文件: %LATEST%
echo [2/3] 复制到项目目录...
copy /y "%LATEST%" "D:\CC\test-claude\sentiment_platform\data\涌益咨询日度数据.xlsx" >nul
if %errorlevel% neq 0 (
    echo [错误] 复制失败！
    pause
    exit /b 1
)

echo [3/3] 提交并推送到 GitHub...
cd /d D:\CC\test-claude\sentiment_platform
git add "data\涌益咨询日度数据.xlsx"
git commit -m "更新现货数据 %date%" 2>nul
git push

echo.
echo ============================================================
echo   同步完成！Streamlit Cloud 将自动重新部署。
echo ============================================================
pause
