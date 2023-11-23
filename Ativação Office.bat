@echo off
>nul 2>&1 "%SYSTEMROOT%\system32\icacls.exe" "%SYSTEMROOT%\system32\config\system"
if '%errorlevel%' NEQ '0' (
    echo Erro: Este script precisa de privilegios de administrador.
    echo Por favor, execute o CMD como administrador e tente novamente.
    pause
    exit /b
)

title Ativador Microsoft Office 2019-2021 (Todas as versoes) gratuitamente - ByRD Tech
cls
echo ===================================================================================================
echo #Projeto: Ativando Microsoft Office 2019-2021 (Pacote Office) gratuitamente, sem software adicional
echo ===================================================================================================
echo.
echo #Produtos suportados:
echo - Microsoft Office Standard 2021
echo - Microsoft Office Professional Plus 2021
echo.
echo.

(if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16")
(if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16")

(for /f %%x in ('dir /b ..\root\Licenses16\ProPlus2021VL_KMS*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)

echo.
echo ========================
echo Ativando seu produto...
cscript //nologo slmgr.vbs /ckms >nul
cscript //nologo ospp.vbs /setprt:1688 >nul
cscript //nologo ospp.vbs /unpkey:6F7TH >nul
set i=1
cscript //nologo ospp.vbs /inpkey:FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH >nul||goto notsupported

:skms
if %i% GTR 10 goto allfailed
if %i% EQU 1 set KMS=kms7.MSGuides.com
if %i% EQU 2 set KMS=e8.us.to
if %i% EQU 3 set KMS=e9.us.to
if %i% GTR 3 goto ato
cscript //nologo ospp.vbs /sethst:%KMS% >nul

:ato
echo ========================
echo.

cscript //nologo ospp.vbs /act | find /i "successful" >nul
if %errorlevel%==0 (
    echo Ativacao bem-sucedida! Seu produto foi ativado com o servidor %KMS%.
    echo.
    echo Encerrando em 5...
    timeout /t 1 >nul
    echo Encerrando em 4...
    timeout /t 1 >nul
    echo Encerrando em 3...
    timeout /t 1 >nul
    echo Encerrando em 2...
    timeout /t 1 >nul
    echo Encerrando em 1...
    timeout /t 1 >nul
    goto end
) else (
    echo Falha na ativacao com o servidor %KMS%. Tentando outro servidor...
    echo.
    set /a i+=1
    goto skms
)

:allfailed
echo =====================================================================================
echo.
echo Falha na ativacao! Todos os servidores de ativacao foram tentados e nenhum funcionou.
echo.
echo Pressione a tecla Enter para sair.
echo.
pause >nul
goto halt

:notsupported
echo =====================================================================================
echo.
echo Desculpe, sua versao nao e suportada.
echo.
echo Pressione a tecla Enter para sair.
echo.
pause >nul
goto halt

:halt
pause >nul

:end
