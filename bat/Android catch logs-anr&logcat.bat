@echo off
cls 
::init 

goto begin

:begin
echo on
adb kill-server
adb start-server
@echo off
set root=d:/
set log_doc=_log/
set anr=ANRs
set logcat=logcat
if exist %root%%log_doc% ( echo Ŀ¼%root%%log_doc%�Ѵ��ڣ����贴��) else ( md %root%%log_doc% )
rem if exist %root%%log_doc%%anr% ( echo Ŀ¼%root%%log_doc%%anr%�Ѵ��ڣ����贴��) else ( md %root%%log_doc%%anr% )
d:
cd %root%%log_doc%
md %anr%
md %logcat%
rem echo %root%%log_doc%%anr%
rem start %root%%log_doc%%anr%
rem if exist %root%%log_doc%%logcat% ( echo Ŀ¼%root%%log_doc%%logcat%�Ѵ��ڣ����贴��) else ( md %root%%log_doc%%logcat% )   
::pause  
set des_doc=%root%%log_doc%%anr%
set des_log=%root%%log_doc%%logcat%
echo �����־��������%root%%log_doc%Ŀ¼��

if "%time:~0,2%" lss "10" (set y=%date:~0,4%%date:~5,2%%date:~8,2%%time:~0,1%%time:~3,2%) else (set y=%date:~0,4%%date:~5,2%%date:~8,2%%time:~0,2%%time:~3,2%)
echo ��ǰʱ�䣺%time% %y%

:goto menu

:menu
echo. 
echo ===================================================
echo   ��ѡ����Ҫ������
echo ===================================================
echo   1 Import out Android Logcat  
echo   2 Get Android Bugreport 
echo   3 Monitor and Import out AndroidRuntime log
echo   4 Import out Android ANR docs 
echo   5 show logcat in console
echo   6 show AndroidRuntime monitor
echo   7 exit
echo ===================================================
echo.
CHOICE /C:1234567  /t 1000 /D 5 /M ��ѡ��:
::set /p choice=%input%
if %errorlevel% geq 7 goto Exit
if %errorlevel%==6 goto ShowRunt
if %errorlevel%==5 goto ShowLogcat
if %errorlevel%==4 goto ANRs
if %errorlevel%==3 goto Runtimelog 
if %errorlevel%==2 goto Bugreport 
if %errorlevel%==1 goto Logcat 

:ShowRunt
echo 3 getting runtime log...
adb logcat -v time -s AndroidRuntime 
::start %des_log%
goto menu

:ShowLogcat 
echo 1 logcatting...
adb logcat -v time 
::> %des_log%/%y%_log.txt
goto menu

:ANRs 
echo Pulling anrs... 
adb pull /data/anr %des_doc%\%y%_anr
echo ANR�ѱ��棬Ŀ¼��%des_doc%\%y%_anr
start %des_doc%\%y%_anr
goto menu 

:Runtimelog
echo 3 getting runtime log...
adb logcat -v time -s AndroidRuntime > %des_log%\%y%_RuntimeLog.txt
start %des_log%
goto menu

:Bugreport 
echo 2 Import bugreport... 
adb bugreport > %des_log%\%y%_bugrep.txt
start %des_log%
goto menu 

:Logcat 
echo 1 logcatting...
adb logcat -v time  > %des_log%\%y%_log.txt
::CHOICE /C:C /d c /n >nul
goto menu 

:Exit
echo  Mession completed��Bye-Bye
choice /t 5 /d y /n >nul 
::sleep 10
exit


