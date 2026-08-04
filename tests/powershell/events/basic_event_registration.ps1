# vybe-test: powershell/events/basic_event_registration
$object = New-Object PSObject -Property @{ Value = 0 }
Register-ObjectEvent -InputObject $object -EventName "TestEvent" -Action {
    $object.Value += 1
}
$object.PSObject.TypeNames.Insert(0, 'TestType')
$object | Add-Member -MemberType ScriptProperty -Name Trigger -Value { $this.PSObject.Properties['Value'].Value = 1 }
$object.Trigger
if ($object.Value -ne 1) {
    Write-Host "FAIL: expected 1, got $($object.Value)"
    exit 1
}
Write-Host "PASS"
exit 0
