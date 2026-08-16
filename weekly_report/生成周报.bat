@echo off
chcp 65001 > nul
title 涌益咨询 - 生猪周度报告系统
echo.
echo ========================================
echo   涌益咨询 · 生猪周度数据报告系统
echo ========================================
echo.
echo   数据源: 数据源\涌益生猪项目数据库\
echo      (1) XX方便拉图表的数据库.xlsx    - 日度数据
echo      (2) XX涌益咨询 周度数据.xlsx     - 周度数据
echo      (3) XX涌益咨询 周度图表版.xlsx   - 图表版
echo.
echo   启动后请访问: http://localhost:8051
echo.
echo   报告内容:
echo   一、价格分析 - 四省价格+屠宰量走势图/季节性同比/全国均价/毛白价差
echo   二、供给分析 - 集团散户体重/二育栏舍利用率/散户肥标价差
echo   三、屠宰分析 - 鲜销率/冻品库存/白条头均利润
echo   四、母猪分析 - 淘汰母猪价格/高低胎折扣/二元母猪价格
echo   五、仔猪分析 - 15公斤仔猪/断奶仔猪/销售仔猪利润
echo   六、养殖利润 - 母猪50头以下/5000-10000头/外购仔猪育肥
echo   七、期货与后市展望
echo.
echo   功能:
echo   - 全部季节性同比图(2021-2026年颜色统一)
echo   - AI自动分析, 按板块输出结论(猪价分析师角度)
echo   - 在线编辑分析文本并保存
echo   - 浏览器 Ctrl+P → 另存为PDF
echo.
echo   更新数据: 替换数据源文件夹中的Excel后重启即可
echo   按 Ctrl+C 停止服务器
echo ========================================
echo.

cd /d "%~dp0"
C:\Users\CC\miniconda3\python.exe generate_report.py

pause
