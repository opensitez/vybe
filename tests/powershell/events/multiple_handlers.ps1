# vybe-test: powershell/events/multiple_handlers
$object = New-Object PSObject -Property @{ Sum = 0 }
Register-ObjectEvent -InputObject $object -EventName "TestEvent" -Action { $object.Sum += 1 }
Register-ObjectEvent -InputObject $object -EventName "TestEvent" -Action { $object.Sum += 2 }
$object.PSObject.Properties.Add((New-Object System.Management.Automation.PSNoteProperty('TestEvent', 'trigger')))
if ($object.Sum -ne 3) {
    Write-Host "FAIL: expected 3, got $($object.Sum)"
    exit 1
}
Write-Host "PASS"
exit 0
