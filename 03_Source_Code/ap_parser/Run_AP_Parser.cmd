@echo off
setlocal

REM ===== Resolve folder paths relative to this CMD file =====
set "ROOT=%~dp0"
set "INPUT=%ROOT%input_ap"
set "OUTPUT_FILE=%ROOT%AP_PAYMENT_TEMPLATE_OUTPUT.xlsx"
set "SCRIPT=%ROOT%ap_payment_template_parser_fixed_v4.py"

REM ===== Basic checks =====
if not exist "%SCRIPT%" (
  echo ERROR: Cannot find parser script:
  echo "%SCRIPT%"
  pause
  exit /b 1
)

if not exist "%INPUT%" (
  echo ERROR: Cannot find input folder:
  echo "%INPUT%"
  pause
  exit /b 1
)

REM ===== Payment-date filter (the batch day) =====
REM The source files are running lists - earlier batches' rows stay in
REM place. Only the rows whose PAYMENT DATE equals the run day belong to
REM this batch, so the parser keeps just that day and drops the old data.
REM Press Enter for today; type a date (YYYY-MM-DD) for another day; type
REM ALL to keep every row like the old behavior.
echo.
set "BATCHDATE="
set /p BATCHDATE="Payment date to keep [Enter = today / YYYY-MM-DD / ALL]: "
if "%BATCHDATE%"=="" set "BATCHDATE=today"

REM ===== Run the active AP parser =====
echo.
echo Running AP parser...
echo Input : "%INPUT%"
echo Output: "%OUTPUT_FILE%"
echo Date  : "%BATCHDATE%"
echo.

py -3 "%SCRIPT%" --input "%INPUT%" --output "%OUTPUT_FILE%" --date "%BATCHDATE%"

if errorlevel 1 (
  echo.
  echo Parser failed. Common cause: "%OUTPUT_FILE%" is currently open in Excel - close it and run this again.
) else (
  echo.
  echo Parser completed successfully. Output: "%OUTPUT_FILE%"
)

pause
endlocal
