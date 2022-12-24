@echo off
taskkill /im chromedriver.exe /f
taskkill /im chrome.exe /f
taskkill /im firefox.exe /f
taskkill /im geckodriver.exe /f
cd /D %temp%
for /d %%D in (*) do rd /s /q "%%D"
del /f /q *