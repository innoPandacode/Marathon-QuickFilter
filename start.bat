@echo off
REM ====================================================
REM start.bat – 每次啟動專案時只要呼叫這個 Batch 檔即可
REM 如果尚未建立 venv，就自動建立並安裝所有套件
REM 如果已經有 venv，則直接啟動並執行 Streamlit
REM ====================================================

REM 先把 Code Page 切到 UTF-8，避免顯示亂碼
chcp 65001 >nul

REM 1) 檢查是否已經有虛擬環境 (venv\Scripts\python.exe)
IF NOT EXIST "venv\Scripts\python.exe" (
    echo.
    echo ================================================
    echo   尚未偵測到「venv」虛擬環境，正在建立中…
    echo ================================================
    echo.
    python -m venv venv
    IF ERRORLEVEL 1 (
        echo ❌ 無法建立虛擬環境，請確認 Python 是否已安裝且已加入 PATH。
        pause
        exit /b 1
    )

    echo.
    echo ================================================
    echo   啟動剛剛建立的虛擬環境…
    echo ================================================
    echo.
    call venv\Scripts\activate.bat
    IF ERRORLEVEL 1 (
        echo ❌ 啟動虛擬環境失敗，請檢查 venv\Scripts\activate.bat 是否存在。
        pause
        exit /b 1
    )

    echo.
    echo ================================================
    echo   更新 pip 並安裝所需套件…
    echo ================================================
    echo.
    python -m pip install --upgrade pip

    REM 下面這行請確保與您的專案需求一致
    pip install streamlit pandas requests beautifulsoup4 reportlab tabulate

    echo.
    echo ================================================
    echo   套件安裝完成！
    echo ================================================
    echo.
) ELSE (
    echo.
    echo ================================================
    echo   已偵測到虛擬環境 venv，跳過建立與安裝步驟
    echo ================================================
    echo.
    call venv\Scripts\activate.bat
    IF ERRORLEVEL 1 (
        echo ❌ 啟動虛擬環境失敗，請檢查 venv\Scripts\activate.bat 是否存在。
        pause
        exit /b 1
    )
)

REM 2) 啟動 Streamlit 應用
echo.
echo ================================================
echo   正在啟動 Streamlit 應用…
echo ================================================
echo.
streamlit run scraper.py

REM 3) Streamlit 停止後才會繼續執行到下面這行
pause
