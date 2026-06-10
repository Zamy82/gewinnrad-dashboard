<#
  update-data.ps1  —  Regeneriert die KAMPAGNEN-Aggregate in data.js / data.json
  aus einem frischen Export des All-in-One-Flyer-Sheets.

  Tageswerte (days) bleiben unangetastet — die kommen aus den 9-Uhr-Mails (Power-Automate-Flow).

  Nutzung:
    .\update-data.ps1 -Xlsx "C:\Pfad\zu\Gewinnspiel.xlsx"
      → exportiert die benoetigten Tabs via Excel nach CSV und rechnet neu.
    .\update-data.ps1 -CsvDir "C:\Pfad\zu\csv_ordner"
      → nutzt bereits exportierte CSVs (Gewinnrad_nutzer.csv, Gewinnrad_Rezi.csv).
#>
param(
  [string]$Xlsx,
  [string]$CsvDir,
  [string]$ProjectDir = $PSScriptRoot
)
$ErrorActionPreference = 'Stop'

if (-not $CsvDir) {
  if (-not $Xlsx) { throw "Bitte -Xlsx ODER -CsvDir angeben." }
  $CsvDir = Join-Path $env:TEMP 'gewinnspiel_csv_update'
  New-Item -ItemType Directory -Force -Path $CsvDir | Out-Null
  Write-Host "Exportiere Tabs aus $Xlsx ..."
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false; $excel.DisplayAlerts = $false
  $wb = $excel.Workbooks.Open($Xlsx)
  foreach ($name in 'Gewinnrad_nutzer','Gewinnrad_Rezi') {
    $ws = $wb.Worksheets.Item($name)
    $out = Join-Path $CsvDir "$name.csv"
    try { $ws.SaveAs($out, 62) } catch { $ws.SaveAs($out, 6) }  # 62=CSV UTF-8, 6=CSV
  }
  $wb.Close($false); $excel.Quit()
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}

# --- NUTZER ---
$hn = 'id','Auftrags_id','Order','Gewinn','Sterne','Seite','Completed','Sprache','Land','Bestmail','Created','Updated','SeitenReihenfolge','x1','x2','x3','x4','Gespielt','ZuEnde','NichtZuEnde','DatumWerte','HeuteGespielt','HeuteZuEnde'
$nutzer = Get-Content (Join-Path $CsvDir 'Gewinnrad_nutzer.csv') -Encoding UTF8 | Select-Object -Skip 1 | ConvertFrom-Csv -Header $hn | Where-Object { $_.Created -and $_.Created.Trim() -ne '' }
# --- REZI ---
$rezi = Import-Csv (Join-Path $CsvDir 'Gewinnrad_Rezi.csv') -Encoding UTF8
$foto = $rezi | Where-Object { $_.Bewertungs_Foto -and $_.Bewertungs_Foto.Trim() -ne '' }

$rated = $nutzer | Where-Object { $_.Sterne -match '^[1-5]$' }
$ssum=0; $rated | ForEach-Object { $ssum += [int]$_.Sterne }
$sterne=[ordered]@{}; 5,4,3,2,1 | ForEach-Object { $v=$_; $sterne["$v"]=@($rated | Where-Object {[int]$_.Sterne -eq $v}).Count }
$land=[ordered]@{}; $nutzer | Group-Object Land | Sort-Object Count -Descending | ForEach-Object { if($_.Name -and $_.Name.Trim() -ne ''){ $land[$_.Name]=$_.Count } }
$gewinn=[ordered]@{}; $nutzer | Where-Object { $_.Gewinn -and $_.Gewinn.Trim() -ne '' } | Group-Object Gewinn | Sort-Object Count -Descending | ForEach-Object { $gewinn[$_.Name]=$_.Count }
$plat=[ordered]@{}; $rezi | Group-Object platform | Sort-Object Count -Descending | ForEach-Object { if($_.Name -and $_.Name.Trim() -ne '' -and $_.Count -ge 3){ $plat[$_.Name]=$_.Count } }
$zuEnde=@($nutzer | Where-Object { $_.Seite -eq 'Zu Ende gespielt' }).Count
$completed=@($nutzer | Where-Object { $_.Completed -eq 'TRUE' }).Count

$campaign=[ordered]@{
  teilnehmer=$nutzer.Count; zuEnde=$zuEnde; completed=$completed
  conversionPct=[math]::Round($zuEnde/$nutzer.Count*100,1)
  fotosHochgeladen=$foto.Count
  fotosValidiert=@($foto | Where-Object {[int]$_.Sterne -gt 0}).Count
  bewertungen=$rated.Count; avgSterne=[math]::Round($ssum/$rated.Count,2)
  sterne=$sterne; plattform=$plat; land=$land; gewinn=$gewinn
}

# --- bestehende Tageswerte (days) aus data.json uebernehmen ---
$existing = Get-Content (Join-Path $ProjectDir 'data.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$today = (Get-Date).ToString('yyyy-MM-dd')

$out=[ordered]@{
  meta=[ordered]@{ source='live'; updated=$today; hinweis=$existing.meta.hinweis }
  campaign=$campaign
  days=$existing.days
}
$json = $out | ConvertTo-Json -Depth 6
$enc = New-Object System.Text.UTF8Encoding($false)
[System.IO.File]::WriteAllText((Join-Path $ProjectDir 'data.json'), $json, $enc)
[System.IO.File]::WriteAllText((Join-Path $ProjectDir 'data.js'), "window.DASHBOARD_DATA = $json;", $enc)
Write-Host "OK — Kampagne aktualisiert: Teilnehmer=$($campaign.teilnehmer) zuEnde=$($campaign.zuEnde) Fotos=$($campaign.fotosHochgeladen). Tage unveraendert: $($out.days.Count)."
