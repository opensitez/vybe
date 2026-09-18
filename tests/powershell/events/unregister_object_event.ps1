# vybe-test: powershell/events/unregister_object_event
$object = New-Object PSObject -Property @{ Count = 0 }
$registration = Register-ObjectEvent -InputObject $object -EventName "TestEvent" -Action {
    $object.Count += 1
}
Unregister-Event -SourceIdentifier $registration.Name
$object.PSObject.Properties.Add((New-Object System.Management.Automation.PSNoteProperty('TestEvent', 'trigger')))
if ($object.Count -ne 0) {
    Write-Host "FAIL: expected 0 after unregister, got $($object.Count)"
    exit 1
}
Write-Host "PASS"
exit 0
