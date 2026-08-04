# vybe-test: powershell/event_sources/register_object_event_source
$object = New-Object PSObject -Property @{ Count = 0 }
Register-ObjectEvent -InputObject $object -EventName TestSrc -Action { $Global.Count = 1 }
$object | Add-Member -MemberType ScriptMethod -Name Raise -Value { New-Event -SourceIdentifier TestSrc }
New-Event -SourceIdentifier TestSrc
if ($Global.Count -ne 1) {
    Write-Host "FAIL: expected event source action"
    exit 1
}
Unregister-Event -SourceIdentifier TestSrc -ErrorAction SilentlyContinue
Write-Host "PASS"
exit 0
