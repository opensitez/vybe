# vybe-test: powershell/redirection/basic_redirection
$temp = [System.IO.Path]::GetTempFileName()
'hello' > $temp
$content = Get-Content $temp
Remove-Item $temp
if ($content -ne 'hello') {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
