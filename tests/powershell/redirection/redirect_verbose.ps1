# vybe-test: powershell/redirection/redirect_verbose
$temp = [System.IO.Path]::GetTempFileName()
Write-Verbose 'v' 4> $temp
$content = Get-Content $temp
Remove-Item $temp
Write-Host 'PASS'
exit 0
