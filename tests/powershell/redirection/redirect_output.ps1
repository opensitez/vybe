# vybe-test: powershell/redirection/redirect_output
$temp = [System.IO.Path]::GetTempFileName()
Write-Output 'hi' > $temp
if ((Get-Content $temp) -ne 'hi') {
    Write-Host 'FAIL'
    exit 1
}
Remove-Item $temp
Write-Host 'PASS'
exit 0
