@echo off
chcp 936 >nul
title 涌益咨询 - 生猪周度报告系统
echo.
echo ============================================
echo   涌益咨询 - 生猪周度数据报告系统
echo ============================================
echo.
echo   数据源: 自动读取 D:\CC\Desktop 最新文件
echo     (1) XX涌益咨询 周度数据.xlsx
echo     (2) XX涌益咨询日度数据.xlsx
echo.
echo   启动后请访问: http://localhost:8051
echo.
echo   报告内容:
echo   一、价格分析   二、供给分析   三、屠宰分析
echo   四、母猪分析   五、仔猪分析   六、养殖利润
echo   七、期货与后市展望
echo.
echo   导出PDF: 浏览器 Ctrl+P 另存为PDF
echo   更新数据: 替换桌面Excel后重启即可
echo   按 Ctrl+C 停止服务器
echo ============================================
echo.
cd /d "%~dp0"
C:\Users\CC\miniconda3\python.exe generate_report.py
pause
