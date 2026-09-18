# vybe-test: powershell/extended_type_system/ets_scriptproperty_returning_boolean
# ScriptProperties can evaluate dynamic boolean state
Update-TypeData -TypeName System.Int32 -MemberType ScriptProperty -MemberName IsEven -Value { ($this % 2) -eq 0 } -Force

$evenCheck = (10).IsEven
$oddCheck  = (11).IsEven

if ($evenCheck -ne $true) {
    Write-Host "FAIL: (10).IsEven expected `$true, got: $evenCheck"
    exit 1
}

if ($oddCheck -ne $false) {
    Write-Host "FAIL: (11).IsEven expected `$false, got: $oddCheck"
    exit 1
}

Write-Host "PASS"
exit 0
