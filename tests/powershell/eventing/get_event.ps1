# vybe-test: powershell/eventing/get_event
New-Event -SourceIdentifier QueryEvent -MessageData 'data'
$ev = Get-Event -SourceIdentifier QueryEvent
if ($ev.SourceIdentifier -ne 'QueryEvent') {
    Write-Host "FAIL: expected QueryEvent"
    exit 1
}
Remove-Event -SourceIdentifier QueryEvent
Write-Host "PASS"
exit 0
