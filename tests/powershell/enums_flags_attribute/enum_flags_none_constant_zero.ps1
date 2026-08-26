# vybe-test: powershell/enums_flags_attribute/enum_flags_none_constant_zero
[System.FlagsAttribute()]
enum OptFlags {
    None = 0
    Opt1 = 1
}
$n = [OptFlags]::None
if ($n.value__ -ne 0 -or $n.HasFlag([OptFlags]::Opt1)) {
    Write-Host "FAIL: None constant check failed"
    exit 1
}
Write-Host "PASS"
exit 0
