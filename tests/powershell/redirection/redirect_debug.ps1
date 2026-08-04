# vybe-test: powershell/redirection/redirect_debug
$temp = [System.IO.Path]::GetTempFileName()
Write-Debug 'd' 5> $temp
Remove-Item $temp
Write-Host 'PASS'
exit 0
