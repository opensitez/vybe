# vybe-test: powershell/one_way_remoting/no_wait_output
$a = Start-Job -ScriptBlock { 5 }
if (-not $a.Id) {
    Write-Host "FAIL: expected job id"
    exit 1
}
Write-Host "PASS"
exit 0
