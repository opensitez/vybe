# vybe-test: powershell/fileio/multiline_file_lines
$path = [System.IO.Path]::GetTempFileName()
$lines = @("line one", "line two", "line three")
Set-Content -Path $path -Value $lines
$read = Get-Content -Path $path
if ($read.Count -ne 3)        { Write-Host "FAIL: count";   Remove-Item $path; exit 1 }
if ($read[0] -ne "line one")  { Write-Host "FAIL: line 0";  Remove-Item $path; exit 1 }
if ($read[2] -ne "line three"){ Write-Host "FAIL: line 2";  Remove-Item $path; exit 1 }
Remove-Item $path
Write-Host "PASS"
exit 0
