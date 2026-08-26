# vybe-test: powershell/enums_flags_attribute/enum_flags_hasflag_method_true
[System.FlagsAttribute()]
enum Perms2 {
    Read = 1
    Write = 2
    Execute = 4
}
$combo = [Perms2]::Read -bor [Perms2]::Execute
if (-not $combo.HasFlag([Perms2]::Read) -or -not $combo.HasFlag([Perms2]::Execute)) {
    Write-Host "FAIL: HasFlag positive check failed"
    exit 1
}
Write-Host "PASS"
exit 0
