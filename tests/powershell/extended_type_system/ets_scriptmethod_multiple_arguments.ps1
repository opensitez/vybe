# vybe-test: powershell/extended_type_system/ets_scriptmethod_multiple_arguments
# ScriptMethods can accept multiple positional parameters
Update-TypeData -TypeName System.Int32 -MemberType ScriptMethod -MemberName Clamp -Value {
    param($min, $max)
    if ($this -lt $min) { return $min }
    if ($this -gt $max) { return $max }
    return $this
} -Force

$low   = (-10).Clamp(0, 100)
$mid   = (50).Clamp(0, 100)
$high  = (150).Clamp(0, 100)

if ($low -ne 0 -or $mid -ne 50 -or $high -ne 100) {
    Write-Host "FAIL: Clamp method returned low=$low, mid=$mid, high=$high"
    exit 1
}

Write-Host "PASS"
exit 0
