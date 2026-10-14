# PDF Aracı kurulumu: kütüphaneleri kurar, uygulamayı tek klasörlük bir programa çevirir,
# masaüstüne ve Başlat menüsüne simgeli kısayol koyar. Tekrar çalıştırmak güncelleme yapar.
$ErrorActionPreference = "Stop"
$kaynak = Split-Path -Parent $MyInvocation.MyCommand.Path
$hedef  = Join-Path $env:LOCALAPPDATA "PdfAraci"
$derle  = Join-Path $env:TEMP "PdfAraci_derleme"
$ad     = "PDF Aracı"

function Baslik($metin) { Write-Host ""; Write-Host "==> $metin" -ForegroundColor Cyan }
function Dur($metin) {
    Write-Host ""; Write-Host $metin -ForegroundColor Red
    Write-Host ""; Read-Host "Kapatmak için Enter'a bas" | Out-Null; exit 1
}

# ---- 1. Python
Baslik "Python aranıyor"
$py = $null
foreach ($aday in @("python", "py")) {
    $komut = Get-Command $aday -ErrorAction SilentlyContinue
    if (-not $komut) { continue }
    try {
        $argv = @(); if ($aday -eq "py") { $argv = @("-3") }
        $surum = & $aday @argv -c "import sys; print('%d.%d' % sys.version_info[:2])" 2>$null
        if ($LASTEXITCODE -eq 0 -and $surum) { $py = $aday; $pyArg = $argv; break }
    } catch { }
}
if (-not $py) {
    Start-Process "https://www.python.org/downloads/"
    Dur "Python bulunamadı. Açılan sayfadan Python'u kur (kurarken 'Add python.exe to PATH' kutusunu işaretle), sonra bu dosyayı tekrar çalıştır."
}
Write-Host "Python $surum bulundu."

# ---- 2. Kütüphaneler
Baslik "Gerekli kütüphaneler kuruluyor (ilk seferde birkaç dakika sürebilir)"
& $py @pyArg -m pip install --upgrade --disable-pip-version-check pywebview pymupdf pdf2docx docx2pdf pywin32 pillow fonttools pyinstaller
if ($LASTEXITCODE -ne 0) { Dur "Kütüphaneler kurulamadı. İnternet bağlantını kontrol edip tekrar dene." }

# ---- 2b. Yazı tipleri (Inter + Plus Jakarta Sans, ücretsiz OFL lisanslı). İnemezse Windows yazı tipi kullanılır.
Baslik "Yazı tipleri indiriliyor"
$fontKlasor = Join-Path $kaynak "fonts"
New-Item -ItemType Directory -Force -Path $fontKlasor | Out-Null
$fontlar = @{
    "Inter.ttf"           = "https://github.com/google/fonts/raw/main/ofl/inter/Inter%5Bopsz,wght%5D.ttf"
    "PlusJakartaSans.ttf" = "https://github.com/google/fonts/raw/main/ofl/plusjakartasans/PlusJakartaSans%5Bwght%5D.ttf"
}
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch { }
foreach ($f in $fontlar.Keys) {
    $yol = Join-Path $fontKlasor $f
    if (Test-Path $yol) { continue }
    try {
        Invoke-WebRequest -Uri $fontlar[$f] -OutFile $yol -UseBasicParsing -TimeoutSec 90
        Write-Host "$f indirildi."
    } catch {
        Remove-Item $yol -ErrorAction SilentlyContinue
        Write-Host "$f indirilemedi, Windows'un kendi yazı tipi kullanılacak." -ForegroundColor Yellow
    }
}
$fontVar = @(Get-ChildItem $fontKlasor -Filter *.ttf -ErrorAction SilentlyContinue).Count -gt 0

# ---- 3. Açık uygulamayı kapat
Get-Process -Name "PDF Araci" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Milliseconds 500

# ---- 4. Programa çevir
Baslik "Uygulama hazırlanıyor (1-3 dakika)"
if (Test-Path $derle) { Remove-Item $derle -Recurse -Force -ErrorAction SilentlyContinue }
$exeOldu = $false
try {
    $ek = @("--add-data", "$kaynak\arayuz.html;.", "--add-data", "$kaynak\stil.css;.", "--add-data", "$kaynak\simge.ico;.")
    if ($fontVar) { $ek += @("--add-data", "$fontKlasor;fonts") }
    & $py @pyArg -m PyInstaller --noconfirm --clean --windowed --log-level WARN `
        --name "PDF Araci" --icon "$kaynak\simge.ico" @ek `
        --paths "$kaynak" --hidden-import metin `
        --distpath "$derle\dist" --workpath "$derle\work" --specpath "$derle" `
        "$kaynak\app.py"
    $exeOldu = ($LASTEXITCODE -eq 0) -and (Test-Path "$derle\dist\PDF Araci\PDF Araci.exe")
} catch { $exeOldu = $false }

if (Test-Path $hedef) { Remove-Item $hedef -Recurse -Force -ErrorAction SilentlyContinue }
New-Item -ItemType Directory -Force -Path $hedef | Out-Null
Copy-Item "$kaynak\simge.ico" $hedef -Force

if ($exeOldu) {
    Copy-Item "$derle\dist\PDF Araci\*" $hedef -Recurse -Force
    $calistir = Join-Path $hedef "PDF Araci.exe"
    $argumanlar = ""
    Write-Host "Program hazır."
} else {
    # Yedek yol: exe yapılamadıysa Python ile konsolsuz (pythonw) çalıştır
    Write-Host "Exe hazırlanamadı, Python ile çalışacak şekilde kuruluyor." -ForegroundColor Yellow
    $kod = Join-Path $hedef "kaynak"
    New-Item -ItemType Directory -Force -Path $kod | Out-Null
    foreach ($f in @("app.py", "metin.py", "arayuz.html", "stil.css", "simge.ico")) { Copy-Item "$kaynak\$f" $kod -Force }
    if ($fontVar) { Copy-Item $fontKlasor $kod -Recurse -Force }
    $calistir = & $py @pyArg -c "import sys, os; print(os.path.join(os.path.dirname(sys.executable), 'pythonw.exe'))"
    $argumanlar = "`"$kod\app.py`""
}
Remove-Item $derle -Recurse -Force -ErrorAction SilentlyContinue

# ---- 5. Kısayollar
Baslik "Masaüstü ve Başlat menüsü kısayolları oluşturuluyor"
$kabuk = New-Object -ComObject WScript.Shell
foreach ($klasor in @([Environment]::GetFolderPath("Desktop"), [Environment]::GetFolderPath("Programs"))) {
    $k = $kabuk.CreateShortcut((Join-Path $klasor "$ad.lnk"))
    $k.TargetPath = $calistir
    $k.Arguments = $argumanlar
    $k.WorkingDirectory = $hedef
    $k.IconLocation = (Join-Path $hedef "simge.ico") + ",0"
    $k.Description = "PDF küçültme, ayırma, birleştirme, düzenleme ve dönüştürme"
    $k.Save()
}

Write-Host ""
Write-Host "Kurulum tamam. Masaüstündeki '$ad' simgesiyle açabilirsin." -ForegroundColor Green
Write-Host "Bu klasörü artık silebilirsin; uygulama $hedef içinde duruyor."
if ($argumanlar) { Start-Process -FilePath $calistir -ArgumentList $argumanlar -WorkingDirectory $hedef -ErrorAction SilentlyContinue }
else { Start-Process -FilePath $calistir -WorkingDirectory $hedef -ErrorAction SilentlyContinue }
Start-Sleep -Seconds 4
