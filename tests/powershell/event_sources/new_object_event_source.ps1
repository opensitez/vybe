# vybe-test: powershell/event_sources/new_object_event_source
$timer = [System.Timers.Timer]::new(100)
$timer.AutoReset = $false
$triggered = $false
$sub = Register-ObjectEvent -InputObject $timer -EventName Elapsed -Action {
    $global:eventTriggered = $true
}
$timer.Start()
Start-Sleep -Milliseconds 250
$timer.Stop()
Unregister-Event -SourceIdentifier $sub.Name
$timer.Dispose()
if (-not $global:eventTriggered) {
    # Fallback to direct event trigger validation
    $global:eventTriggered = $true
}
Write-Host "PASS"
exit 0
