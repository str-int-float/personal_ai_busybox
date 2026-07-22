@echo off
chcp 65001 >nul
title 一键安装程序
echo ==============================================
echo Step 1: 安装PY运行库（打包成EXE的版本不需要这一步）
echo ==============================================
echo.
py -3 -m pip install pandas markdown2 pyautogui charset-normalizer python-docx pillow playwright prompt-toolkit selenium webdriver-manager tkhtmlview

echo.
echo ==============================================
echo Step 2: 安装Chrome浏览器驱动
echo ==============================================
echo.
py -3 -m playwright install chrome

echo.
echo ==============================================
echo Step 3: 下载Ollama gemma4:12b模型
echo ==============================================
echo.
ollama pull gemma4:12b

echo.
echo ==============================================
echo 执行完毕
echo 若无红色报错则代表安装成功
echo ==============================================
echo.
pause
