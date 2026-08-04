# vybe-test: powershell/redirection/error_redirection
$temp = [System.IO.Path]::GetTempFileName()
Write-Error 'err' 2> $temp
$content = Get-Content $temp
Remove-Item $temp
if ($content.Length -lt 1) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
