@echo off
chcp 65001 >nul
setlocal
pushd "%~dp0"
if exist ".venv-build\Scripts\python.exe" goto install
py -3 -c "import sys; assert sys.version_info >= (3,10)" >nul 2>&1
if not errorlevel 1 (
  py -3 -m venv .venv-build
  if errorlevel 1 goto fail
  goto install
)
python -c "import sys; assert sys.version_info >= (3,10)" >nul 2>&1
if not errorlevel 1 (
  python -m venv .venv-build
  if errorlevel 1 goto fail
  goto install
)
echo Нужен Python 3.10 или новее для Windows с компонентом Tcl/Tk.
echo Установите Python с https://www.python.org/downloads/windows/
echo При установке отметьте Add python.exe to PATH, затем повторите запуск.
goto fail
:install
echo Устанавливаем библиотеки. Для первой сборки нужен интернет.
".venv-build\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 goto fail
".venv-build\Scripts\python.exe" build_windows.py
if errorlevel 1 goto fail
echo Готово. Сервер и приложение запущены. Оба EXE находятся в папке dist.
start "" explorer.exe "%CD%\dist"
popd
pause
exit /b 0
:fail
echo Сборка не завершена. Смотрите сообщение об ошибке выше.
popd
pause
exit /b 1
