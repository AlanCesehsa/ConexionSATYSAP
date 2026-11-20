@echo off
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo [INFO] Solicitando permisos de administrador...
    powershell -Command "Start-Process '%~f0' -Verb RunAs"
    exit /b
)

echo ========================================================
echo  Instalando Certificado CESEHSA Desarrollos SAP Addons
echo ========================================================
echo.
certutil -addstore -f "TrustedPublisher" "%~dp0CESEHSA_Addons.cer"
certutil -addstore -f "Root" "%~dp0CESEHSA_Addons.cer"

echo.
echo ========================================================
echo  Certificado instalado correctamente con exito!
echo  Ya puedes abrir SAP Business One.
echo ========================================================
pause