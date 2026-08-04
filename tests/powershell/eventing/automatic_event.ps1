# vybe-test: powershell/eventing/automatic_event
Register-EngineEvent -SourceIdentifier AutoEvent -Action { $Global.Fired = $true }
New-Event -SourceIdentifier AutoEvent
if (-not $Global.Fired) {
    Write-Host "FAIL: expected fired"
    exit 1
}
Unregister-Event -SourceIdentifier AutoEvent -ErrorAction SilentlyContinue
Write-Host "PASS"
exit 0
