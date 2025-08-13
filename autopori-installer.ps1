# autopori-installer.ps1 – Asenna AutoPori, rekisteröi suoritettavaksi kirjautuessa ja suorita heti
# Suositus: Aja asennusvalvojan oikeuksin (elevated) jotta scheduled task -rekisteröinti voi pyytää korkeampia oikeuksia.

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

# Luo optimoitu päivittäinen ajoskripti
$dailyRunnerTemplate = @'
# autopori.ps1 – Hae sarjakuva ja aseta taustakuvaksi (optimized)

$feedUrl  = "{0}"
$savePath = "{1}"

# Precompile .NET assembly only once
if (-not ("Wallpaper" -as [type])) {{
    Add-Type @"  
    using System.Runtime.InteropServices;
    public class Wallpaper {{
        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni);
    }}
"@
}}

# Fetch comic with timeout
try {{
    $response = Invoke-WebRequest -Uri $feedUrl -UseBasicParsing -TimeoutSec 15 -ErrorAction Stop
    [xml]$xml = $response.Content
}} catch {{
    Write-Host "RSS-syötteen hakeminen epäonnistui: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}}

# Etsi kuvan URL
$description = $xml.rss.channel.item[0].description.'#cdata-section'
if ($description -match '<img.+?src="(.+?)"') {{
    $imageUrl = $matches[1] -replace '\?.*$',''
}} else {{
    Write-Host "Kuvan URL-osoitetta ei löytynyt." -ForegroundColor Red
    exit 1
}}

# Lataa kuva
try {{
    Invoke-WebRequest -Uri $imageUrl -OutFile $savePath -TimeoutSec 15 -ErrorAction Stop
}} catch {{
    Write-Host "Kuvan lataaminen epäonnistui: $imageUrl" -ForegroundColor Red
    exit 1
}}

# Tarkista onko taustakuva jo sama
$currentWallpaper = [Microsoft.Win32.Registry]::GetValue(
    "HKCU:\Control Panel\Desktop", 
    "Wallpaper", 
    ""
)

if ($currentWallpaper -ne $savePath) {{
    # Aseta vain jos kuva on muuttunut
    Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name WallpaperStyle -Value "6"
    Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name TileWallpaper   -Value "0"
    [Wallpaper]::SystemParametersInfo(20, 0, $savePath, 3)
}}
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

# Poista Scheduled Task (AutoPori)
try {{
    if (Get-ScheduledTask -TaskName "AutoPori" -ErrorAction SilentlyContinue) {{
        Unregister-ScheduledTask -TaskName "AutoPori" -Confirm:$false
        Write-Host "Scheduled task 'AutoPori' poistettu." -ForegroundColor Green
    }}
}} catch {{
    # fallback: schtasks
    schtasks /Delete /TN "AutoPori" /F | Out-Null
}}

# Poista HKCU Run -avain, jos asetettu
try {{
    if (Get-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run" -Name "AutoPori" -ErrorAction SilentlyContinue) {{
        Remove-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run" -Name "AutoPori" -ErrorAction SilentlyContinue
        Write-Host "Run-avain poistettu." -ForegroundColor Green
    }}
}} catch {{}}

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
$shortcut.Arguments       = "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$mainScriptPath`""
$shortcut.WorkingDirectory = $targetFolder
$shortcut.Save()

# Luo optimoitu tehtävä XML-määrittely (nopeampi käynnistys)
$taskXml = @"
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.4" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Description>AutoPori: fetch daily comic and set wallpaper</Description>
  </RegistrationInfo>
  <Triggers>
    <LogonTrigger>
      <Enabled>true</Enabled>
    </LogonTrigger>
  </Triggers>
  <Principals>
    <Principal id="Author">
      <UserId>$($env:USERDOMAIN)\$($env:USERNAME)</UserId>
      <LogonType>InteractiveToken</LogonType>
      <RunLevel>LeastPrivilege</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
    <AllowHardTerminate>true</AllowHardTerminate>
    <StartWhenAvailable>true</StartWhenAvailable>
    <RunOnlyIfNetworkAvailable>true</RunOnlyIfNetworkAvailable>
    <IdleSettings>
      <StopOnIdleEnd>true</StopOnIdleEnd>
      <RestartOnIdle>false</RestartOnIdle>
    </IdleSettings>
    <AllowStartOnDemand>true</AllowStartOnDemand>
    <Enabled>true</Enabled>
    <Hidden>false</Hidden>
    <RunOnlyIfIdle>false</RunOnlyIfIdle>
    <WakeToRun>false</WakeToRun>
    <ExecutionTimeLimit>PT1M</ExecutionTimeLimit>
    <Priority>5</Priority>
  </Settings>
  <Actions Context="Author">
    <Exec>
      <Command>powershell.exe</Command>
      <Arguments>-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "$mainScriptPath"</Arguments>
    </Exec>
  </Actions>
</Task>
"@

$taskXmlPath = Join-Path $env:TEMP "AutoPoriTask.xml"
$taskXml | Set-Content -Path $taskXmlPath -Encoding Unicode

# Rekisteröi tehtävä käyttäen optimoitua XML-määritystä
try {
    schtasks.exe /Create /XML "$taskXmlPath" /TN "AutoPori" /F
    Write-Host "Scheduled task luotu optimoiduilla asetuksilla." -ForegroundColor Green
    Remove-Item $taskXmlPath -Force -ErrorAction SilentlyContinue
} catch {
    Write-Host "ScheduledTask XML -rekisteröinti epäonnistui, kokeillaan schtasks-fallbackia..." -ForegroundColor Yellow
    try {
        $arguments = @(
            "/Create",
            "/TN", "AutoPori",
            "/TR", "`"powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$mainScriptPath`"`"",
            "/SC", "ONLOGON",
            "/RL", "LIMITED",
            "/F"
        )
        Start-Process -FilePath "schtasks.exe" -ArgumentList $arguments -NoNewWindow -Wait -ErrorAction Stop
        Write-Host "Scheduled task luotu schtasksilla." -ForegroundColor Green
    } catch {
        Write-Host "Scheduled taskin luonti epäonnistui." -ForegroundColor Red
    }
}

# Lisää HKCU Run -avain fallbackiksi
try {
    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run" -Name "AutoPori" -Value "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$mainScriptPath`""
    Write-Host "HKCU Run -avain asetettu fallbackiksi." -ForegroundColor Green
} catch {
    Write-Host "HKCU Run -avaimen asettaminen epäonnistui." -ForegroundColor Yellow
}

# Suorita skripti heti asennuksen jälkeen
try {
    Start-Process -FilePath "powershell.exe" -ArgumentList "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$mainScriptPath`"" -WindowStyle Hidden
    Write-Host "Autopori käynnistetty välittömästi." -ForegroundColor Cyan
} catch {
    Write-Host "Autoporin välitön käynnistys epäonnistui." -ForegroundColor Red
}

Write-Host "`nSarjakuva asennettu onnistuneesti!" -ForegroundColor Cyan
Write-Host "Valittu syöte: $feedUrl"
Write-Host "Taustakuva päivittyy kirjautumisen yhteydessä (Scheduled Task)."  
Write-Host "Ohjelman poisto: suorita työpöydällä oleva 'remove-autopori.ps1'." -ForegroundColor Yellow