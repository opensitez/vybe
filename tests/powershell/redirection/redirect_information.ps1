# vybe-test: powershell/redirection/redirect_information
$temp = [System.IO.Path]::GetTempFileName()
Write-Information 'info' 6> $temp
Remove-Item $temp
Write-Host 'PASS'
exit 0
