# vybe-test: powershell/command_invocation/command_with_redirection
$temp = [System.IO.Path]::GetTempFileName()
Write-Output 'PASS' > $temp
if ((Get-Content $temp) -eq 'PASS') { Remove-Item $temp; Write-Host 'PASS'; exit 0 }
Remove-Item $temp
Write-Host 'FAIL'
exit 1
