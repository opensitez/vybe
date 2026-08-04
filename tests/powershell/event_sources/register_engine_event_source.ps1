# vybe-test: powershell/event_sources/register_engine_event_source
Register-EngineEvent -SourceIdentifier EngineSource -Action { $Global.SourceFired = $true }
New-Event -SourceIdentifier EngineSource
if (-not $Global.SourceFired) {
    Write-Host "FAIL: expected engine source fired"
    exit 1
}
Unregister-Event -SourceIdentifier EngineSource -ErrorAction SilentlyContinue
Write-Host "PASS"
exit 0
