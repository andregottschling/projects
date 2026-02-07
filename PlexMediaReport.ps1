# =====================================================
# KONFIGURATION
# =====================================================

$verzeichnisse = @(
    "\\mynas\Serien5TB",
    "\\mynas\HorrorServ",
    "\\mynas\VideoServer 2"
)

$videoExtensions = @(
    ".mkv", ".mp4", ".avi", ".mov", ".wmv", ".m4v",
    ".m2ts", ".ts", ".mpg"
)

$basePath   = "\\mynas\Public\PlexCont"
$csvPath    = Join-Path $basePath "csv"
$secureCred = "\\mynas\secure\smtp.xml"

$smtpServer = "send.one.com"
$smtpPort   = 587
$mailFrom   = "cloud@gottschling.it"
$mailTo     = "info@gottschling.it"

$datum = Get-Date -Format "yyyy-MM-dd"
$zeit  = Get-Date -Format "HHmmss"
$monat = Get-Date -Format "yyyy-MM"

$logFile = Join-Path $basePath "Backup_$datum.txt"

# =====================================================
# VORBEREITUNG
# =====================================================

$ErrorActionPreference = "Continue"
New-Item -ItemType Directory -Path $csvPath -Force | Out-Null

$currentFileCsvTimestamp = Join-Path $csvPath "PLEX_FILES_${datum}_$zeit.csv"
$currentFileCsvMonthly   = Join-Path $csvPath "PLEX_FILES_$monat.csv"
$currentSummaryCsv       = Join-Path $csvPath "PLEX_SUMMARY_$monat.csv"
$duplicateCsv            = Join-Path $csvPath "PLEX_DUPLICATES_${datum}_$zeit.csv"
$trendCsv                = Join-Path $csvPath "PLEX_TREND.csv"

$lastMonth   = (Get-Date).AddMonths(-1).ToString("yyyy-MM")
$lastFileCsv = Join-Path $csvPath "PLEX_FILES_$lastMonth.csv"

# =====================================================
# START LOG
# =====================================================

Start-Transcript -Path $logFile

# =====================================================
# ERFASSUNG
# =====================================================

$fileList = @()
$result   = @()
$totalMB  = 0

foreach ($dir in $verzeichnisse) {

    $files = Get-ChildItem -Path $dir -Recurse -File -ErrorAction SilentlyContinue |
             Where-Object { $videoExtensions -contains $_.Extension.ToLower() }

    foreach ($file in $files) {

        $sizeMB = [math]::Round($file.Length / 1MB, 2)
        $group  = Split-Path $file.DirectoryName -Leaf

        $fileList += [PSCustomObject]@{
            Basisverzeichnis = $dir
            Gruppe           = $group
            Datei            = $file.Name
            Vollpfad         = $file.FullName
            Größe_MB         = $sizeMB
        }
    }

    $folderMB = if ($files.Count -gt 0) {
        [math]::Round(($files | Measure-Object Length -Sum).Sum / 1MB, 2)
    } else { 0 }

    $totalMB += $folderMB

    $result += [PSCustomObject]@{
        Verzeichnis = $dir
        Größe_MB    = $folderMB
    }
}

# =====================================================
# DUBLETTEN ERKENNEN
# =====================================================

$duplicates = $fileList |
Group-Object Datei, Größe_MB |
Where-Object { $_.Count -gt 1 } |
ForEach-Object {
    $_.Group
}

if ($duplicates.Count -gt 0) {
    $duplicates |
    Select Basisverzeichnis, Gruppe, Datei, Größe_MB, Vollpfad |
    Export-Csv $duplicateCsv -NoTypeInformation -Encoding UTF8
}

# =====================================================
# MONATS-TREND
# =====================================================

$trendEntry = [PSCustomObject]@{
    Monat     = $monat
    Gesamt_MB = [math]::Round($totalMB, 2)
}

if (Test-Path $trendCsv) {
    $trendData = Import-Csv $trendCsv
    if ($trendData.Monat -notcontains $monat) {
        $trendData += $trendEntry
        $trendData | Export-Csv $trendCsv -NoTypeInformation -Encoding UTF8
    }
} else {
    $trendEntry | Export-Csv $trendCsv -NoTypeInformation -Encoding UTF8
}

$trendData = Import-Csv $trendCsv

# ASCII-Sparkline
$spark = ($trendData | ForEach-Object {
    $bars = [math]::Round($_.Gesamt_MB / 100000)
    ('█' * [math]::Min($bars, 20))
}) -join '<br>'

# =====================================================
# CSV EXPORT
# =====================================================

$fileList |
Export-Csv $currentFileCsvTimestamp -NoTypeInformation -Encoding UTF8

$fileList |
Export-Csv $currentFileCsvMonthly -NoTypeInformation -Encoding UTF8

$result |
Export-Csv $currentSummaryCsv -NoTypeInformation -Encoding UTF8

# =====================================================
# HTML MAIL
# =====================================================

$dirRows = foreach ($r in ($result | Sort-Object Größe_MB -Descending)) {

    $bg = if ($r.Größe_MB -eq 0) { " style='background:#f8d7da'" } else { "" }

    "<tr$bg>
        <td>$($r.Verzeichnis)</td>
        <td align='right'>$($r.Größe_MB)</td>
     </tr>"
}

$htmlBody = @"
<html>
<body style="font-family:Arial;font-size:13px">
<h3>PLEX Medienreport ($datum)</h3>

<h4>Verzeichnisse</h4>
<table border="1" cellpadding="5">
<tr><th>Verzeichnis</th><th>MB</th></tr>
$dirRows
</table>

<p>
<b>Dubletten:</b> $($duplicates.Count)<br>
</p>

<h4>Monats-Trend</h4>
$spark

</body>
</html>
"@

# =====================================================
# MAILVERSAND
# =====================================================

if (-not (Test-Path $secureCred)) {
    throw "SMTP Credential-Datei fehlt: $secureCred"
}

$message = New-Object System.Net.Mail.MailMessage
$message.From = $mailFrom
$message.To.Add($mailTo)
$message.Subject = "PLEX Medienreport $datum"
$message.Body = $htmlBody
$message.IsBodyHtml = $true
$message.Attachments.Add($currentFileCsvTimestamp)

if ($duplicates.Count -gt 0) {
    $message.Attachments.Add($duplicateCsv)
}

$smtp = New-Object System.Net.Mail.SmtpClient($smtpServer, $smtpPort)
$smtp.EnableSsl = $true
$smtp.Credentials = Import-Clixml $secureCred
$smtp.Send($message)

Stop-Transcript
