title Office 2024 Install Tool
@echo off
setlocal EnableDelayedExpansion

:: Yonetici yetkisi kontrolu ve yukseltme
>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"
if '%errorlevel%' NEQ '0' (
    echo Administrator rights required...
    goto UACPrompt
) else ( goto gotAdmin )

:UACPrompt
    echo Set UAC = CreateObject^("Shell.Application"^) > "%temp%\getadmin.vbs"
    echo UAC.ShellExecute "%~s0", "", "", "runas", 1 >> "%temp%\getadmin.vbs"
    "%temp%\getadmin.vbs"
    exit /B

:gotAdmin
    if exist "%temp%\getadmin.vbs" ( del "%temp%\getadmin.vbs" )
    pushd "%CD%"
    CD /D "%~dp0"

:: Dil secimi
echo ====================================
echo 1. Turkce
echo 2. English
echo 3. Deutsch
echo ====================================
set /p lang_choice="Dil/Language/Sprache (1-3): " || set "lang_choice=1"

:: Dil secimine gore degiskenleri ayarla
if "%lang_choice%"=="1" goto Turkish
if "%lang_choice%"=="2" goto English
if "%lang_choice%"=="3" goto German
goto Turkish

:Turkish
set "MSG_TITLE=Office Kurulum Yapilandirma Araci"
set "MSG_SYSTEM_TYPE=Tespit edilen sistem:"
set "MSG_BIT=bit Windows"
set "MSG_CHANNEL=Kanal (PerpetualVL2024, CurrentPreview, MonthlyEnterprise) [PerpetualVL2024]: "
set "MSG_PRODUCT=urun ID (ProPlus2024Volume, StandardVL2024) [ProPlus2024Volume]: "
set "MSG_LANGUAGES=Yuklemek istediginiz diller:"
set "MSG_MAIN_LANG=Ana dil (tr-tr, en-us, vb.) [tr-tr]: "
set "MSG_SECOND_LANG=Ikinci dil (birakabilirsiniz) [en-us]: "
set "MSG_APPS=Uygulamalar (Dahil etmek icin 'E', haric tutmak icin 'H' yazin)"
set "MSG_WORD=Word (E/H) [E]: "
set "MSG_EXCEL=Excel (E/H) [E]: "
set "MSG_POWERPOINT=PowerPoint (E/H) [E]: "
set "MSG_ACCESS=Access (E/H) [H]: "
set "MSG_OUTLOOK=Outlook (E/H) [H]: "
set "MSG_PUBLISHER=Publisher (E/H) [H]: "
set "MSG_ONENOTE=OneNote (E/H) [H]: "
set "MSG_ONEDRIVE=OneDrive (E/H) [H]: "
set "MSG_LYNC=Teams/Lync (E/H) [H]: "
set "MSG_AUTOACTIVATE=Otomatik etkinlestirme (1-Evet, 0-Hayir) [0]: "
set "MSG_DISPLAYLEVEL=Kurulum arayuzu (None, Full) [None]: "
set "MSG_EULA=EULA kabul (TRUE, FALSE) [TRUE]: "
set "MSG_UPDATES=Guncellemeleri etkinlestir (TRUE, FALSE) [TRUE]: "
set "MSG_CONFIG_CREATED=Yapilandirma dosyasi olusturuldu:"
set "MSG_DOWNLOADING=Setup.exe indiriliyor..."
set "MSG_DOWNLOAD_COMPLETE=Indirme tamamlandi."
set "MSG_SETUP_STARTING=Office kurulumu baslatiliyor..."
set "MSG_SETUP_STARTED=Kurulum baslatildi."
set "MSG_CLEANING=Gecici dosyalar temizleniyor..."
set "MSG_CLEANING_COMPLETE=Temizleme tamamlandi."
set "MSG_COMPLETED=Islem tamamlandi."
set "YES_CHAR=E"
set "DEFAULT_MAIN_LANG=tr-tr"
goto Setup

:English
set "MSG_TITLE=Office Setup Configuration Tool"
set "MSG_SYSTEM_TYPE=Detected system:"
set "MSG_BIT=bit Windows"
set "MSG_CHANNEL=Channel (PerpetualVL2024, CurrentPreview, MonthlyEnterprise) [PerpetualVL2024]: "
set "MSG_PRODUCT=Product ID (ProPlus2024Volume, StandardVL2024) [ProPlus2024Volume]: "
set "MSG_LANGUAGES=Languages to install:"
set "MSG_MAIN_LANG=Primary language (en-us, de-de, etc.) [en-us]: "
set "MSG_SECOND_LANG=Secondary language (leave empty if not needed) []: "
set "MSG_APPS=Applications (Enter 'Y' to include, 'N' to exclude)"
set "MSG_WORD=Word (Y/N) [Y]: "
set "MSG_EXCEL=Excel (Y/N) [Y]: "
set "MSG_POWERPOINT=PowerPoint (Y/N) [Y]: "
set "MSG_ACCESS=Access (Y/N) [N]: "
set "MSG_OUTLOOK=Outlook (Y/N) [N]: "
set "MSG_PUBLISHER=Publisher (Y/N) [N]: "
set "MSG_ONENOTE=OneNote (Y/N) [N]: "
set "MSG_ONEDRIVE=OneDrive (Y/N) [N]: "
set "MSG_LYNC=Teams/Lync (Y/N) [N]: "
set "MSG_AUTOACTIVATE=Auto activate (1-Yes, 0-No) [0]: "
set "MSG_DISPLAYLEVEL=Display level (None, Full) [None]: "
set "MSG_EULA=Accept EULA (TRUE, FALSE) [TRUE]: "
set "MSG_UPDATES=Enable updates (TRUE, FALSE) [TRUE]: "
set "MSG_CONFIG_CREATED=Configuration file created:"
set "MSG_DOWNLOADING=Downloading Setup.exe..."
set "MSG_DOWNLOAD_COMPLETE=Download completed."
set "MSG_SETUP_STARTING=Starting Office setup..."
set "MSG_SETUP_STARTED=Setup initiated."
set "MSG_CLEANING=Cleaning temporary files..."
set "MSG_CLEANING_COMPLETE=Cleanup completed."
set "MSG_COMPLETED=Process completed."
set "YES_CHAR=Y"
set "DEFAULT_MAIN_LANG=en-us"
goto Setup

:German
set "MSG_TITLE=Office-Setup-Konfigurationstool"
set "MSG_SYSTEM_TYPE=Erkanntes System:"
set "MSG_BIT=bit Windows"
set "MSG_CHANNEL=Kanal (PerpetualVL2024, CurrentPreview, MonthlyEnterprise) [PerpetualVL2024]: "
set "MSG_PRODUCT=Produkt-ID (ProPlus2024Volume, StandardVL2024) [ProPlus2024Volume]: "
set "MSG_LANGUAGES=Zu installierende Sprachen:"
set "MSG_MAIN_LANG=Hauptsprache (de-de, en-us, usw.) [de-de]: "
set "MSG_SECOND_LANG=Zweitsprache (leer lassen, wenn nicht benotigt) []: "
set "MSG_APPS=Anwendungen (Geben Sie 'J' zum Einschließen, 'N' zum Ausschließen ein)"
set "MSG_WORD=Word (J/N) [J]: "
set "MSG_EXCEL=Excel (J/N) [J]: "
set "MSG_POWERPOINT=PowerPoint (J/N) [J]: "
set "MSG_ACCESS=Access (J/N) [N]: "
set "MSG_OUTLOOK=Outlook (J/N) [N]: "
set "MSG_PUBLISHER=Publisher (J/N) [N]: "
set "MSG_ONENOTE=OneNote (J/N) [N]: "
set "MSG_ONEDRIVE=OneDrive (J/N) [N]: "
set "MSG_LYNC=Teams/Lync (J/N) [N]: "
set "MSG_AUTOACTIVATE=Automatisch aktivieren (1-Ja, 0-Nein) [0]: "
set "MSG_DISPLAYLEVEL=Anzeigeebene (None, Full) [None]: "
set "MSG_EULA=EULA akzeptieren (TRUE, FALSE) [TRUE]: "
set "MSG_UPDATES=Updates aktivieren (TRUE, FALSE) [TRUE]: "
set "MSG_CONFIG_CREATED=Konfigurationsdatei erstellt:"
set "MSG_DOWNLOADING=Setup.exe wird heruntergeladen..."
set "MSG_DOWNLOAD_COMPLETE=Download abgeschlossen."
set "MSG_SETUP_STARTING=Office-Setup wird gestartet..."
set "MSG_SETUP_STARTED=Setup gestartet."
set "MSG_CLEANING=Temporäre Dateien werden bereinigt..."
set "MSG_CLEANING_COMPLETE=Bereinigung abgeschlossen."
set "MSG_COMPLETED=Vorgang abgeschlossen."
set "YES_CHAR=J"
set "DEFAULT_MAIN_LANG=de-de"
goto Setup

:Setup
echo %MSG_TITLE%
echo --------------------------------
echo.

:: 32 bit veya 64 bit Windows tespiti
if exist "%ProgramFiles(x86)%" (
    set "office_edition=64"
) else (
    set "office_edition=32"
)
echo %MSG_SYSTEM_TYPE% %office_edition% %MSG_BIT%

:: Yapilandirma dosyasi icin bir ID olustur
set "config_id=%random%%random%-%random%-%random%-%random%-%random%%random%"

:: Kullanici girdileri
set /p channel="%MSG_CHANNEL%" || set "channel=PerpetualVL2024"
set /p product_id="%MSG_PRODUCT%" || set "product_id=ProPlus2024Volume"

echo.
echo %MSG_LANGUAGES%
set /p main_lang="%MSG_MAIN_LANG%" || set "main_lang=%DEFAULT_MAIN_LANG%"
set /p second_lang="%MSG_SECOND_LANG%" || set "second_lang="

echo.
echo %MSG_APPS%
set /p inc_word="%MSG_WORD%" || set "inc_word=%YES_CHAR%"
set /p inc_excel="%MSG_EXCEL%" || set "inc_excel=%YES_CHAR%"
set /p inc_powerpoint="%MSG_POWERPOINT%" || set "inc_powerpoint=%YES_CHAR%"
set /p inc_access="%MSG_ACCESS%" || set "inc_access=N"
set /p inc_outlook="%MSG_OUTLOOK%" || set "inc_outlook=N"
set /p inc_publisher="%MSG_PUBLISHER%" || set "inc_publisher=N"
set /p inc_onenote="%MSG_ONENOTE%" || set "inc_onenote=N"
set /p inc_onedrive="%MSG_ONEDRIVE%" || set "inc_onedrive=N"
set /p inc_lync="%MSG_LYNC%" || set "inc_lync=N"

echo.
set /p auto_activate="%MSG_AUTOACTIVATE%" || set "auto_activate=0"
set /p display_level="%MSG_DISPLAYLEVEL%" || set "display_level=None"
set /p accept_eula="%MSG_EULA%" || set "accept_eula=TRUE"
set /p enable_updates="%MSG_UPDATES%" || set "enable_updates=TRUE"

:: Gecici dizinde calisalim
cd /d "%temp%"

:: Yapilandirma dosyasini olustur
echo ^<Configuration ID="%config_id%"^> > config.xml
echo   ^<Add OfficeClientEdition="%office_edition%" Channel="%channel%"^> >> config.xml
echo     ^<Product ID="%product_id%"^> >> config.xml
echo       ^<Language ID="%main_lang%"/^> >> config.xml

if not "%second_lang%"=="" (
    echo       ^<Language ID="%second_lang%"/^> >> config.xml
)

if /i "%inc_access%" NEQ "%YES_CHAR%" echo       ^<ExcludeApp ID="Access"/^> >> config.xml
if /i "%inc_lync%" NEQ "%YES_CHAR%" echo       ^<ExcludeApp ID="Lync"/^> >> config.xml
if /i "%inc_onedrive%" NEQ "%YES_CHAR%" echo       ^<ExcludeApp ID="OneDrive"/^> >> config.xml
if /i "%inc_onenote%" NEQ "%YES_CHAR%" echo       ^<ExcludeApp ID="OneNote"/^> >> config.xml
if /i "%inc_outlook%" NEQ "%YES_CHAR%" echo       ^<ExcludeApp ID="Outlook"/^> >> config.xml
if /i "%inc_publisher%" NEQ "%YES_CHAR%" echo       ^<ExcludeApp ID="Publisher"/^> >> config.xml
if /i "%inc_word%" NEQ "%YES_CHAR%" echo       ^<ExcludeApp ID="Word"/^> >> config.xml
if /i "%inc_excel%" NEQ "%YES_CHAR%" echo       ^<ExcludeApp ID="Excel"/^> >> config.xml
if /i "%inc_powerpoint%" NEQ "%YES_CHAR%" echo       ^<ExcludeApp ID="PowerPoint"/^> >> config.xml

echo     ^</Product^> >> config.xml
echo   ^</Add^> >> config.xml
echo   ^<Property Name="SharedComputerLicensing" Value="0"/^> >> config.xml
echo   ^<Property Name="FORCEAPPSHUTDOWN" Value="FALSE"/^> >> config.xml
echo   ^<Property Name="DeviceBasedLicensing" Value="0"/^> >> config.xml
echo   ^<Property Name="SCLCacheOverride" Value="0"/^> >> config.xml
echo   ^<Property Name="AUTOACTIVATE" Value="%auto_activate%"/^> >> config.xml
echo   ^<Updates Enabled="%enable_updates%"/^> >> config.xml
echo   ^<RemoveMSI/^> >> config.xml
echo   ^<Display Level="%display_level%" AcceptEULA="%accept_eula%"/^> >> config.xml
echo ^</Configuration^> >> config.xml

echo.
echo %MSG_CONFIG_CREATED% %temp%\config.xml
echo.

:: Setup.exe'yi indir
echo %MSG_DOWNLOADING%
if exist "%temp%\setup.exe" del "%temp%\setup.exe"
curl -o setup.exe https://officecdn.microsoft.com/pr/wsus/setup.exe
echo %MSG_DOWNLOAD_COMPLETE%

:: Kurulumu baslat
echo %MSG_SETUP_STARTING%
"%temp%\setup.exe" /configure "%temp%\config.xml"
echo %MSG_SETUP_STARTED%

:: Temizlik
echo %MSG_CLEANING%
timeout /t 5 /nobreak > nul
del "%temp%\config.xml" /f /q
del "%temp%\setup.exe" /f /q
echo %MSG_CLEANING_COMPLETE%

echo.
echo %MSG_COMPLETED%
pause