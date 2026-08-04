# vybe-test: powershell/one_way_remoting/queue_command_async
$job = Start-Job -ScriptBlock { 1 + 2 }
if (-not $job) {
    Write-Host "FAIL: expected job object"
    exit 1
}
Write-Host "PASS"
exit 0
