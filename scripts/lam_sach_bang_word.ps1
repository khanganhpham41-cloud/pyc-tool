param([string]$ListFile)
# Moi dong: <duong dan .docx>|<so trang dung>  (so trang phieu + 1 trang Thoa thuan)
$lines = Get-Content -LiteralPath $ListFile -Encoding UTF8 | Where-Object { $_.Trim() -ne "" }
if (-not $lines) { exit 0 }
try { $word = New-Object -ComObject Word.Application } catch {
  Write-Host "Khong mo duoc Word tu dong - hay mo tung file bang Word de kiem tra truoc khi gui." ; exit 0 }
$word.Visible = $false
$word.DisplayAlerts = 0
$bad = 0
foreach ($ln in $lines) {
  $parts = $ln.Split('|')
  $full = $parts[0]
  $expect = 0
  if ($parts.Length -gt 1) { [int]::TryParse($parts[1], [ref]$expect) | Out-Null }
  $name = Split-Path $full -Leaf
  try {
    $doc = $word.Documents.Open($full, $false, $false, $false)
    $doc.Repaginate()
    $pages = $doc.ComputeStatistics(2)
    $doc.SaveAs2($full, 16)
    $doc.Close($false)
    if ($expect -gt 0 -and $pages -ne $expect) {
      $bad++
      Write-Host ("!!! TRAN TRANG: {0} -> Word dem {1} trang, dung ra phai {2}. Giam 'So mau/trang' trong NHAP_MAU.xlsx roi chay lai." -f $name, $pages, $expect) -ForegroundColor Red
    } else {
      Write-Host ("OK  {0}  ({1} trang, dung so trang)" -f $name, $pages) -ForegroundColor Green
    }
  } catch {
    $bad++
    Write-Host ("!!! Word khong mo/luu duoc {0}: {1}" -f $name, $_.Exception.Message) -ForegroundColor Red
  }
}
$word.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
if ($bad -gt 0) { Write-Host ""; Write-Host "CO $bad FILE CAN KIEM TRA LAI (xem dong do o tren)." -ForegroundColor Red }
