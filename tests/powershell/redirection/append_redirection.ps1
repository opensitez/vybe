# vybe-test: powershell/redirection/append_redirection
$temp = [System.IO.Path]::GetTempFileName()
'one' > $temp
'two' >> $temp
$content = Get-Content $temp
Remove-Item $temp
if ($content.Length -ne 2) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
