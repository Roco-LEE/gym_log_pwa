@echo off
chcp 65001 >nul
title 헬스일지 서버
cd /d "%~dp0"
echo.
echo  [헬스일지 서버]  창을 닫으면 서버가 꺼집니다.
echo  --test 를 붙이면 테스트 엑셀에 동기화:  서버켜기.bat --test
echo.
python serve.py %*
echo.
pause
