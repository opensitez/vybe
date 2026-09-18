# vybe-test: powershell/extended_type_system/ets_overwrite_with_force_succeeds
# -Force allows replacing an existing ETS dynamic property definition
class ForceTarget { [string]$V = "Initial" }

Update-TypeData -TypeName ForceTarget -MemberType ScriptProperty -MemberName DynVal -Value { "Version 1" } -Force
$obj = [ForceTarget]::new()

if ($obj.DynVal -ne "Version 1") {
    Write-Host "FAIL: initial value mismatch, got: '$($obj.DynVal)'"
    exit 1
}

Update-TypeData -TypeName ForceTarget -MemberType ScriptProperty -MemberName DynVal -Value { "Version 2" } -Force

if ($obj.DynVal -ne "Version 2") {
    Write-Host "FAIL: overwritten value with -Force failed, got: '$($obj.DynVal)'"
    exit 1
}

Write-Host "PASS"
exit 0
