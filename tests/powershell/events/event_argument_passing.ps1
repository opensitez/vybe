# vybe-test: powershell/events/event_argument_passing
$object = New-Object PSObject -Property @{ Value = 0 }
Register-ObjectEvent -InputObject $object -EventName "TestEvent" -Action {
    param($sender, $eventArgs)
    $object.Value = $eventArgs.NewValue
}
$eventArgs = [PSCustomObject]@{ NewValue = 5 }
$object.PSObject.Properties.Add((New-Object System.Management.Automation.PSNoteProperty('TestEvent', $eventArgs)))
if ($object.Value -ne 0) {
    Write-Host "FAIL: event handler should not have run yet"
    exit 1
}
Write-Host "PASS"
exit 0
