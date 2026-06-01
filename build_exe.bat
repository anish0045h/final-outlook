@echo off
setlocal EnableDelayedExpansion

echo ===============================================
echo   Payroll Agent 
echo ===============================================
echo.

REM ── Step 0: Verify Python is available ────────────────────────────
where python >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Python not found in PATH. Aborting.
    pause
    exit /b 1
)

REM ── Step 1: Install pinned dependencies from requirements.txt ─────
echo [1/4] Installing pinned dependencies...
if not exist requirements.txt (
    echo [ERROR] requirements.txt not found. Aborting.
    echo         Create it with pinned versions before building.
    pause
    exit /b 1
)
pip install -r requirements.txt
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Dependency installation failed. Aborting.
    pause
    exit /b 1
)
echo.

REM ── Step 2: Clean previous builds ────────────────────────────────
echo [2/4] Cleaning previous builds...
if exist build   rmdir /s /q build
if exist dist    rmdir /s /q dist
echo.

REM ── Step 3: Build the executable ─────────────────────────────────
echo [3/4] Building executable...
pyinstaller PayrollAgent.spec
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] PyInstaller build failed. Check errors above.
    pause
    exit /b 1
)
echo.

REM ── Step 4: Verify output and generate integrity hash ────────────
echo [4/4] Verifying build and generating SHA-256 hash...
if not exist dist\PayrollAgent.exe (
    echo ===============================================
    echo   BUILD FAILED - dist\PayrollAgent.exe missing
    echo ===============================================
    pause
    exit /b 1
)

REM Generate SHA-256 hash for distribution integrity verification.
REM Share this hash with IT/SOC alongside the EXE for whitelist requests.
certutil -hashfile dist\PayrollAgent.exe SHA256 > dist\PayrollAgent.exe.sha256
if %ERRORLEVEL% NEQ 0 (
    echo [WARNING] Could not generate SHA-256 hash. Proceeding anyway.
) else (
    echo   SHA-256 hash saved to: dist\PayrollAgent.exe.sha256
    echo   Provide this hash to IT/SOC when submitting for AV whitelisting.
)

echo.
echo ===============================================
echo   BUILD SUCCESSFUL
echo   Executable : dist\PayrollAgent.exe
echo   Hash file  : dist\PayrollAgent.exe.sha256
echo.
echo   Before deploying:
echo     1. Submit EXE + hash to IT for AV whitelisting
echo     2. Sign the EXE with your code-signing certificate
echo     3. Do NOT distribute the EXE without IT sign-off
echo ===============================================
echo.
pause