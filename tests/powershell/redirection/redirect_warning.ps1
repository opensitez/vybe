# vybe-test: powershell/redirection/redirect_warning
$temp = [System.IO.Path]::GetTempFileName()
Write-Warning 'warn' 3> $temp
$content = Get-Content $temp
Remove-Item $temp
if ($content.Length -lt 1) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
