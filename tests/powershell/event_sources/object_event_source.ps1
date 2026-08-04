# vybe-test: powershell/event_sources/object_event_source
$obj = New-Object PSObject
Register-ObjectEvent -InputObject $obj -EventName ObjEvent -Action { $Global.Raised = $true }
New-Event -SourceIdentifier ObjEvent
if (-not $Global.Raised) {
    Write-Host "FAIL: expected object event source"
    exit 1
}
Unregister-Event -SourceIdentifier ObjEvent -ErrorAction SilentlyContinue
Write-Host "PASS"
exit 0
