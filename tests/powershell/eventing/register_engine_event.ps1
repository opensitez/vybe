# vybe-test: powershell/eventing/register_engine_event
$event = Register-EngineEvent -SourceIdentifier TestEngineEvent -Action { $Global:EventFired = $true }
New-Event -SourceIdentifier TestEngineEvent
if (-not $Global:EventFired) {
    Write-Host "FAIL: expected event fired"
    Unregister-Event -SourceIdentifier TestEngineEvent -ErrorAction SilentlyContinue
    exit 1
}
Unregister-Event -SourceIdentifier TestEngineEvent -ErrorAction SilentlyContinue
Write-Host "PASS"
exit 0
