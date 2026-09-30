@echo off
REM ============================================================
REM  Yoklama Uygulamasi - Windows EXE Derleme
REM ============================================================
REM  Bu dosyayi, INTERNETI OLAN ve PYTHON kurulu bir Windows
REM  bilgisayarda CIFT TIKLAYARAK calistirin (kurum bilgisayari
REM  olmak zorunda degil - internet olan herhangi bir bilgisayar
REM  olabilir, ornegin evinizdeki bilgisayar).
REM
REM  Bu islem SADECE BIR KERE yapilir. Sonucunda olusan
REM  "dist\Yoklama.exe" dosyasini, internet gerektirmeden,
REM  kurumdaki TUM bilgisayarlara kopyalayip calistirabilirsiniz.
REM
REM  "logo.ico" dosyasi da yoklama_app.py ile AYNI KLASORDE ise,
REM  olusan EXE'nin simgesi (ikonu) bu logo olur. logo.ico yoksa
REM  EXE yine de sorunsuz olusur, sadece varsayilan simgeyle.
REM ============================================================

setlocal
cd /d "%~dp0"

echo ============================================
echo  Yoklama Uygulamasi - EXE Derleme
echo ============================================
echo.

where python >nul 2>nul
if errorlevel 1 (
    echo HATA: Bu bilgisayarda Python bulunamadi.
    echo.
    echo Lutfen once https://www.python.org/downloads/ adresinden Python indirip kurun.
    echo Kurulum ekraninda "Add python.exe to PATH" kutucugunu MUTLAKA isaretleyin.
    echo Kurulumdan sonra bu dosyayi tekrar calistirin.
    echo.
    pause
    exit /b 1
)

echo Python bulundu. Gerekli paketler kuruluyor / guncelleniyor...
python -m pip install --upgrade pip
python -m pip install --upgrade pyinstaller openpyxl tkcalendar

if errorlevel 1 (
    echo.
    echo HATA: Gerekli paketler kurulamadi. Internet baglantinizi kontrol edin.
    pause
    exit /b 1
)

echo.
echo EXE dosyasi olusturuluyor, lutfen bekleyin (1-2 dakika surebilir)...
if exist "logo.ico" (
    python -m PyInstaller --noconfirm --onefile --windowed --name Yoklama --icon=logo.ico yoklama_app.py
) else (
    echo NOT: "logo.ico" bu klasorde bulunamadi, EXE varsayilan simgeyle olusturulacak.
    python -m PyInstaller --noconfirm --onefile --windowed --name Yoklama yoklama_app.py
)

if errorlevel 1 (
    echo.
    echo HATA: Derleme basarisiz oldu. Yukaridaki hata mesajina bakin.
    pause
    exit /b 1
)

echo.
echo ============================================
echo  TAMAMLANDI!
echo  EXE dosyaniz: %~dp0dist\Yoklama.exe
echo.
echo  Bu "Yoklama.exe" dosyasini kurumdaki tum bilgisayarlara
echo  kopyalayabilirsiniz. Kopyaladiginiz bilgisayarlarda Python
echo  veya internet GEREKMEZ.
echo ============================================
pause