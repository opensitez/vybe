# vybe-test: powershell/event_sources/cleanup_source
New-Event -SourceIdentifier CleanupSource
Remove-Event -SourceIdentifier CleanupSource
$event = Get-Event -SourceIdentifier CleanupSource -ErrorAction SilentlyContinue
if ($event) {
    Write-Host "FAIL: expected removed source events"
    exit 1
}
Write-Host "PASS"
exit 0
