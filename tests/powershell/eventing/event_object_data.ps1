# vybe-test: powershell/eventing/event_object_data
New-Event -SourceIdentifier ObjectEvent -MessageData @{ Value = 2 }
$ev = Get-Event -SourceIdentifier ObjectEvent
if ($ev.MessageData.Value -ne 2) {
    Write-Host "FAIL: expected object data value 2"
    exit 1
}
Remove-Event -SourceIdentifier ObjectEvent
Write-Host "PASS"
exit 0
