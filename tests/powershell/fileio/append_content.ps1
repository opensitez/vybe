# vybe-test: powershell/fileio/append_content
$path = [System.IO.Path]::GetTempFileName()
Set-Content    -Path $path -Value "first"
Add-Content    -Path $path -Value "second"
$lines = Get-Content -Path $path
if ($lines.Count -ne 2)     { Write-Host "FAIL: count $($lines.Count)"; Remove-Item $path; exit 1 }
if ($lines[1] -ne "second") { Write-Host "FAIL: second line";           Remove-Item $path; exit 1 }
Remove-Item $path
Write-Host "PASS"
exit 0
