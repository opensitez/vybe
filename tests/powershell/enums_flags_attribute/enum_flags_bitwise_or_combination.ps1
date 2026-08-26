# vybe-test: powershell/enums_flags_attribute/enum_flags_bitwise_or_combination
[System.FlagsAttribute()]
enum Permissions {
    None = 0
    Read = 1
    Write = 2
    Execute = 4
}
$rw = [Permissions]::Read -bor [Permissions]::Write
if ($rw.value__ -ne 3) {
    Write-Host "FAIL: Bitwise OR flags combination failed, got $($rw.value__)"
    exit 1
}
Write-Host "PASS"
exit 0
