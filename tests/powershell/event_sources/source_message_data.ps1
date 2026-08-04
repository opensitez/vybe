# vybe-test: powershell/event_sources/source_message_data
New-Event -SourceIdentifier DataSource -MessageData 'hello'
$event = Get-Event -SourceIdentifier DataSource
if ($event.MessageData -ne 'hello') {
    Write-Host "FAIL: expected message data hello"
    exit 1
}
Remove-Event -SourceIdentifier DataSource
Write-Host "PASS"
exit 0
