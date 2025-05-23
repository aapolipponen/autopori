# autopori-installer.ps1 – Asenna AutoPori ja valitse sarjakuvasyöte

$targetFolder     = "$env:APPDATA\autopori"
$mainScriptPath   = Join-Path $targetFolder "autopori.ps1"
$startupFolder    = [Environment]::GetFolderPath("Startup")
$shortcutPath     = Join-Path $startupFolder "LaunchAutopori.lnk"
$desktopPath      = [Environment]::GetFolderPath("Desktop")
$removeScriptPath = Join-Path $desktopPath "remove-autopori.ps1"
$savePath         = "$([Environment]::GetFolderPath('MyPictures'))\comic_today.jpg"

# Kysy käyttäjältä haluamaansa RSS-syötettä
Write-Host "Valitse haluamasi RSS-syöte:" -ForegroundColor Yellow
# 1) Fingerpori – Helsingin Sanomat (Aapon suositus) (HS.fi, RSS: darkball.net)
Write-Host "  1) Fingerpori – Helsingin Sanomat " -NoNewline
Write-Host "(Aapon suositus) " -ForegroundColor Cyan -NoNewline
Write-Host "(HS.fi, RSS: darkball.net) " -ForegroundColor DarkGray
# 2) Fingerpori – Kaleva (kaleva.fi, RSS: kimmo.suominen.com)
Write-Host "  2) Fingerpori – Kaleva " -NoNewline
Write-Host "(kaleva.fi, RSS: kimmo.suominen.com)" -ForegroundColor DarkGray
# 3) Fingerpori värillisenä – (fingerpori.org, RSS: darkball.net)
Write-Host "  3) Fingerpori värillisenä – " -NoNewline
Write-Host "(fingerpori.org, RSS: darkball.net)" -ForegroundColor DarkGray
# 4) Viivi ja Wagner – Helsingin Sanomat (HS.fi, RSS: darkball.net)
Write-Host "  4) Viivi ja Wagner – Helsingin Sanomat " -NoNewline
Write-Host "(HS.fi, RSS: darkball.net)" -ForegroundColor DarkGray
# 5) Kamala Luonto – Keski-Suomen Lehtimedia (ksml.fi, RSS: darkball.net)
Write-Host "  5) Kamala Luonto – Keski-Suomen Lehtimedia " -NoNewline
Write-Host "(ksml.fi, RSS: darkball.net)" -ForegroundColor DarkGray

$choice = Read-Host "Anna valintasi numerona (1–5)"

switch ($choice) {
    "1" { $feedUrl = "https://darkball.net/fingerpori/" }
    "2" { $feedUrl = "https://kimmo.suominen.com/stuff/fingerpori-k.xml" }
    "3" { $feedUrl = "https://darkball.net/fingerporiorg/" }
    "4" { $feedUrl = "https://darkball.net/viivijawagner/" }
    "5" { $feedUrl = "https://darkball.net/kamalaluonto/" }
    default {
        Write-Host "Virheellinen valinta. Ohjelma suljetaan." -ForegroundColor Red
        exit 1
    }
}

# Luo hakemisto, jos sitä ei ole
if (-not (Test-Path $targetFolder)) {
    New-Item -ItemType Directory -Path $targetFolder | Out-Null
}

# Luo päivittäinen ajoskripti (autopori.ps1)
$dailyRunnerTemplate = @'
# autopori.ps1 – Hae sarjakuva ja aseta taustakuvaksi

$feedUrl  = "{0}"
$savePath = "{1}"

try {{
    $response = Invoke-WebRequest -Uri $feedUrl -UseBasicParsing
    [xml]$xml = $response.Content
}} catch {{
    Write-Host "RSS-syötteen hakeminen tai jäsentäminen epäonnistui: $feedUrl" -ForegroundColor Red
    exit 1
}}

$description = $xml.rss.channel.item[0].description.'#cdata-section'
if ($description -match '<img.+?src="(.+?)"') {{
    $imageUrl = $matches[1] -replace '\?.*$',''
}} else {{
    Write-Host "Kuvan URL-osoitetta ei löytynyt." -ForegroundColor Red
    exit 1
}}

Invoke-WebRequest -Uri $imageUrl -OutFile $savePath

# Aseta taustakuva
Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name WallpaperStyle -Value "6"
Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name TileWallpaper   -Value "0"

Add-Type @"  
using System.Runtime.InteropServices;
public class Wallpaper {{
    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni);
}}
"@

[Wallpaper]::SystemParametersInfo(20, 0, $savePath, 3)
'@ -f $feedUrl, $savePath

$dailyRunnerTemplate | Set-Content -Encoding UTF8 -Path $mainScriptPath

# Luo poistoskripti työpöydälle
$removalTemplate = @'
# remove-autopori.ps1 – Poista AutoPori ja palauta taustakuva

Add-Type @"  
using System.Runtime.InteropServices;
public class Wallpaper {{
    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni);
}}
"@

$SPI_SETDESKWALLPAPER = 0x0014
$SPIF_UPDATEINIFILE  = 0x01
$SPIF_SENDCHANGE     = 0x02

# Poista käynnistyskuvake
$shortcut = Join-Path ([Environment]::GetFolderPath("Startup")) "LaunchAutopori.lnk"
if (Test-Path $shortcut) {{
    Remove-Item $shortcut -Force
    Write-Host "Poistettu käynnistyskuvake." -ForegroundColor Green
}}

# Poista tallennettu kuva
$imagePath = "{1}"
if (Test-Path $imagePath) {{
    Remove-Item $imagePath -Force
    Write-Host "Poistettu tallennettu sarjakuva." -ForegroundColor Green
}}

# Poista AppData-kansio
$autoporiFolder = "{2}"
if (Test-Path $autoporiFolder) {{
    Remove-Item $autoporiFolder -Recurse -Force
    Write-Host "Poistettu AutoPori-kansio." -ForegroundColor Green
}}

# Palauta taustakuva mustaksi
Add-Type -AssemblyName System.Drawing
$tempWall = "$env:TEMP\black_wallpaper.bmp"
$bmp = New-Object Drawing.Bitmap 1,1
$bmp.SetPixel(0,0,[Drawing.Color]::Black)
$bmp.Save($tempWall,[System.Drawing.Imaging.ImageFormat]::Bmp)
$bmp.Dispose()

[Wallpaper]::SystemParametersInfo(
    $SPI_SETDESKWALLPAPER,
    0,
    $tempWall,
    $SPIF_UPDATEINIFILE -bor $SPIF_SENDCHANGE
)

Write-Host "AutoPori poistettu ja taustakuva palautettu." -ForegroundColor Green

# Itsepoisto
$scriptPath = $MyInvocation.MyCommand.Definition
if (Test-Path $scriptPath) {{
    Remove-Item $scriptPath -Force
    Write-Host "Poistoskripti poistettu itsestään." -ForegroundColor Green
}}
'@ -f $feedUrl, $savePath, $targetFolder

$removalTemplate | Set-Content -Encoding UTF8 -Path $removeScriptPath

# Luo käynnistyskuvake päivittäiselle ajolle
$WScriptShell = New-Object -ComObject WScript.Shell
$shortcut      = $WScriptShell.CreateShortcut($shortcutPath)
$shortcut.TargetPath      = "powershell.exe"
$shortcut.Arguments       = "-ExecutionPolicy Bypass -WindowStyle Hidden -File `"$mainScriptPath`""
$shortcut.WorkingDirectory = $targetFolder
$shortcut.Save()

# Pyöritä autopori.ps1
Start-Process powershell -ArgumentList "-ExecutionPolicy Bypass -WindowStyle Hidden -File `"$mainScriptPath`""

Write-Host "`nSarjakuva asennettu onnistuneesti!" -ForegroundColor Cyan
Write-Host "Valittu syöte: $feedUrl"
Write-Host "Taustakuva päivittyy kirjautumisen yhteydessä."  
Write-Host "Ohjelman poisto: suorita työpöydällä oleva 'remove-autopori.ps1'." -ForegroundColor Yellow