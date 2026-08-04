# vybe-test: powershell/redirection/redirect_error_from_command
$temp = [System.IO.Path]::GetTempFileName()
Get-Command UnknownCmd 2> $temp | Out-Null
$content = Get-Content $temp
Remove-Item $temp
if ($content.Length -lt 1) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
