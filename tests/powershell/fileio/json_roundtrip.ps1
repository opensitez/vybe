# vybe-test: powershell/fileio/json_roundtrip
$path = [System.IO.Path]::GetTempFileName()
$obj = [PSCustomObject]@{ Name = "Alice"; Score = 99 }
$obj | ConvertTo-Json | Set-Content -Path $path
$loaded = Get-Content -Path $path | ConvertFrom-Json
if ($loaded.Name  -ne "Alice") { Write-Host "FAIL: Name";  Remove-Item $path; exit 1 }
if ($loaded.Score -ne 99)      { Write-Host "FAIL: Score"; Remove-Item $path; exit 1 }
Remove-Item $path
Write-Host "PASS"
exit 0
