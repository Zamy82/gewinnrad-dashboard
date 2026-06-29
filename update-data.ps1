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

# Plattform 'Sportstech.de' zaehlt als Trustpilot (so wie in den Tagesberichten)
function Norm($p){ if($p -eq 'Sportstech.de'){ 'Trustpilot' } else { $p } }
function FoldPlat($rows, $min){
  $acc=@{}
  foreach($r in $rows){ $n=$r.platform; if(-not $n -or $n.Trim() -eq ''){ continue }; $n=(Norm $n); if($acc.ContainsKey($n)){ $acc[$n]++ } else { $acc[$n]=1 } }
  $o=[ordered]@{}; foreach($kv in ($acc.GetEnumerator() | Sort-Object Value -Descending)){ if($kv.Value -ge $min){ $o[$kv.Key]=$kv.Value } }
  $o
}
function ProductList($rows){
  $list=@()
  foreach($g in ($rows | Group-Object Geraet | Sort-Object Count -Descending)){
    if(-not $g.Name -or $g.Name.Trim() -eq ''){ continue }
    $list += [ordered]@{ product=$g.Name; total=$g.Count; platforms=(FoldPlat $g.Group 1) }
  }
  ,$list
}

$rated = $nutzer | Where-Object { $_.Sterne -match '^[1-5]$' }
$ssum=0; $rated | ForEach-Object { $ssum += [int]$_.Sterne }
$sterne=[ordered]@{}; 5,4,3,2,1 | ForEach-Object { $v=$_; $sterne["$v"]=@($rated | Where-Object {[int]$_.Sterne -eq $v}).Count }
$land=[ordered]@{}; $nutzer | Group-Object Land | Sort-Object Count -Descending | ForEach-Object { if($_.Name -and $_.Name.Trim() -ne ''){ $land[$_.Name]=$_.Count } }
$gewinn=[ordered]@{}; $nutzer | Where-Object { $_.Gewinn -and $_.Gewinn.Trim() -ne '' } | Group-Object Gewinn | Sort-Object Count -Descending | ForEach-Object { $gewinn[$_.Name]=$_.Count }
$plat = FoldPlat $rezi 3
$zuEnde=@($nutzer | Where-Object { $_.Seite -eq 'Zu Ende gespielt' }).Count
$completed=@($nutzer | Where-Object { $_.Completed -eq 'TRUE' }).Count

$campaign=[ordered]@{
  teilnehmer=$nutzer.Count; zuEnde=$zuEnde; completed=$completed
  conversionPct=[math]::Round($zuEnde/$nutzer.Count*100,1)
  fotosHochgeladen=$foto.Count
  fotosValidiert=@($foto | Where-Object {[int]$_.Sterne -gt 0}).Count
  bewertungen=$rated.Count; avgSterne=[math]::Round($ssum/$rated.Count,2)
  sterne=$sterne; plattform=$plat; land=$land; gewinn=$gewinn
  reviewedProducts=(ProductList $foto)
}

# --- Monats-Verteilungen (Plattform/Land/Gewinn/Sterne pro Monat, aus dem Sheet) ---
$monthly=[ordered]@{}
foreach($g in ($nutzer | Group-Object { $_.Created.Substring(0,7) } | Sort-Object Name)){
  $m=$g.Name
  $rs = $g.Group | Where-Object { $_.Sterne -match '^[1-5]$' }
  $st=[ordered]@{}; 5,4,3,2,1 | ForEach-Object { $v=$_; $st["$v"]=@($rs | Where-Object {[int]$_.Sterne -eq $v}).Count }
  $ld=[ordered]@{}; $g.Group | Group-Object Land | Sort-Object Count -Descending | ForEach-Object { if($_.Name -and $_.Name.Trim() -ne ''){ $ld[$_.Name]=$_.Count } }
  $gw=[ordered]@{}; $g.Group | Where-Object { $_.Gewinn -and $_.Gewinn.Trim() -ne '' } | Group-Object Gewinn | Sort-Object Count -Descending | ForEach-Object { $gw[$_.Name]=$_.Count }
  $monthly[$m]=[ordered]@{ teilnehmer=$g.Count; sterne=$st; land=$ld; gewinn=$gw; plattform=[ordered]@{} }
}
foreach($g in ($rezi | Where-Object { $_.Zeit_des_Spiels -and $_.Zeit_des_Spiels.Length -ge 7 } | Group-Object { $_.Zeit_des_Spiels.Substring(0,7) })){
  $m=$g.Name
  if(-not $monthly.Contains($m)){ $monthly[$m]=[ordered]@{teilnehmer=0;sterne=[ordered]@{};land=[ordered]@{};gewinn=[ordered]@{};plattform=[ordered]@{}} }
  $monthly[$m].plattform = FoldPlat $g.Group 2
  $mfoto = $g.Group | Where-Object { $_.Bewertungs_Foto -and $_.Bewertungs_Foto.Trim() -ne '' }
  $monthly[$m].reviewedProducts = (ProductList $mfoto)
}

# --- bestehende Tageswerte (days) aus data.js uebernehmen (kommen aus den Mails) ---
# WICHTIG: data.js ist maßgeblich — der 9-Uhr-Tagesbot committet NUR data.js, nicht data.json.
# Daher days/meta IMMER aus data.js lesen, sonst werden neue Tageswerte ueberschrieben.
$djsRaw = Get-Content (Join-Path $ProjectDir 'data.js') -Raw -Encoding UTF8
$existing = ($djsRaw -replace '^\s*window\.DASHBOARD_DATA\s*=\s*','' -replace ';\s*$','') | ConvertFrom-Json
$today = (Get-Date).ToString('yyyy-MM-dd')

$out=[ordered]@{
  meta=[ordered]@{ source='live'; updated=$today; hinweis=$existing.meta.hinweis }
  campaign=$campaign
  monthly=$monthly
  days=$existing.days
}
$json = $out | ConvertTo-Json -Depth 12
$enc = New-Object System.Text.UTF8Encoding($false)
[System.IO.File]::WriteAllText((Join-Path $ProjectDir 'data.json'), $json, $enc)
[System.IO.File]::WriteAllText((Join-Path $ProjectDir 'data.js'), "window.DASHBOARD_DATA = $json;", $enc)
Write-Host "OK — Kampagne: Teilnehmer=$($campaign.teilnehmer) zuEnde=$($campaign.zuEnde) Fotos=$($campaign.fotosHochgeladen). Monats-Verteilungen: $($monthly.Count) Monate. Tage (aus Mails) unveraendert: $($out.days.Count)."
