# vybe-test: powershell/event_sources/new_object_event_source
$obj = New-Object PSObject
Register-ObjectEvent -InputObject $obj -EventName TestSource -Action { $Global.Fired = $true }
New-Event -SourceIdentifier TestSource
if (-not $Global.Fired) {
    Write-Host "FAIL: expected event source fired"
    exit 1
}
Unregister-Event -SourceIdentifier TestSource -ErrorAction SilentlyContinue
Write-Host "PASS"
exit 0
