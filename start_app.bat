@echo off
setlocal

set "PYTHON_EXE=E:\anaconda\python.exe"

if not exist "%PYTHON_EXE%" (
    echo [ERROR] Python not found: %PYTHON_EXE%
    pause
    exit /b 1
)

if not exist "app.py" (
    echo [ERROR] app.py not found in current directory.
    pause
    exit /b 1
)

if not exist "courses.parquet" (
    if exist "test.xlsx" (
        echo [INFO] courses.parquet not found. Converting from test.xlsx...
        "%PYTHON_EXE%" convert_data.py
        if errorlevel 1 (
            echo [ERROR] Failed to generate courses.parquet
            pause
            exit /b 1
        )
    ) else (
        echo [WARN] courses.parquet and test.xlsx are both missing.
        echo [WARN] The app can still start, but data will be unavailable until you provide a file.
    )
)

echo [INFO] Starting Streamlit app on http://localhost:8501 ...
start "" http://localhost:8501
"%PYTHON_EXE%" -m streamlit run app.py --server.headless true

if errorlevel 1 (
    echo.
    echo [ERROR] Streamlit failed to start.
    pause
    exit /b 1
)
