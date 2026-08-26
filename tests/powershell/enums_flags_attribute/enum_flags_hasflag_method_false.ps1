# vybe-test: powershell/enums_flags_attribute/enum_flags_hasflag_method_false
[System.FlagsAttribute()]
enum Perms3 {
    Read = 1
    Write = 2
    Execute = 4
}
$combo = [Perms3]::Read -bor [Perms3]::Execute
if ($combo.HasFlag([Perms3]::Write)) {
    Write-Host "FAIL: HasFlag negative check failed"
    exit 1
}
Write-Host "PASS"
exit 0
