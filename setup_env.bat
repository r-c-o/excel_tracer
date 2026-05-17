@echo off
REM ============================================================
REM  Excel Tracer - Conda Environment Setup
REM ============================================================

SET CONDA_EXE=C:\Users\ryanc\miniconda3\Scripts\conda.exe
SET ENV_NAME=excel_tracer

echo.
echo [1/3] Checking for existing conda environment "%ENV_NAME%"...
%CONDA_EXE% env list | findstr /C:"%ENV_NAME%" >nul 2>&1
IF %ERRORLEVEL% EQU 0 (
    echo      Found existing environment. Removing it...
    %CONDA_EXE% env remove -n %ENV_NAME% -y
)

echo.
echo [2/3] Creating conda environment from environment.yml...
%CONDA_EXE% env create -f environment.yml
IF %ERRORLEVEL% NEQ 0 (
    echo ERROR: Failed to create conda environment.
    pause
    exit /b 1
)

echo.
echo [3/3] Verifying installed packages...
%CONDA_EXE% run -n %ENV_NAME% python -c "import openpyxl; import networkx; import pyparsing; import yaml; print('All packages verified OK.')"
IF %ERRORLEVEL% NEQ 0 (
    echo ERROR: Package verification failed.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo  Setup complete!
echo  To activate: conda activate %ENV_NAME%
echo  To run: conda run -n %ENV_NAME% python run_all.py workbook.xlsx
echo ============================================================
echo.
pause
