# vybe-test: powershell/event_sources/event_source_metadata
New-Event -SourceIdentifier MetadataSource -MessageData @{ Name = 'x' }
$ev = Get-Event -SourceIdentifier MetadataSource
if ($ev.MessageData.Name -ne 'x') {
    Write-Host "FAIL: expected metadata x"
    exit 1
}
Remove-Event -SourceIdentifier MetadataSource
Write-Host "PASS"
exit 0
