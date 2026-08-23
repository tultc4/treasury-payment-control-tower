@echo off
setlocal
cd /d "%~dp0"

if "%~1"=="" (
  echo ================================================================
  echo BANK STATEMENT TRANSACTION PARSER
  echo ================================================================
  echo Keo folder chua so phu tha vao file BAT nay,
  echo hoac nhap duong dan folder ben duoi.
  echo.
  set /p INPUT_PATH=Nhap duong dan file/folder so phu: 
) else (
  set "INPUT_PATH=%~1"
)

if "%INPUT_PATH%"=="" (
  echo Khong co duong dan input.
  pause
  exit /b 1
)

if not exist "%INPUT_PATH%" (
  echo Khong tim thay: %INPUT_PATH%
  pause
  exit /b 1
)

if not exist "output" mkdir "output"

python "bank_statement_transaction_parser.py" --input "%INPUT_PATH%" --output-dir "%~dp0output" --config "%~dp0parser_config.json" --verbose
set EXIT_CODE=%ERRORLEVEL%

echo.
if %EXIT_CODE% EQU 0 (
  echo HOAN TAT.
  echo Output: %~dp0output\bank_transaction_detail.xlsx
  echo Keo file nay vao Run_4_Import_Bank_Statement.bat cua dashboard.
) else (
  echo Parser ket thuc voi ma loi %EXIT_CODE%.
  echo Mo output\bank_transaction_detail.xlsx va sheet PARSER_LOG neu file da duoc tao.
)

echo.
pause
exit /b %EXIT_CODE%
