# vybe-test: powershell/enums_flags_attribute/enum_flags_bitwise_xor_toggle_flag
[System.FlagsAttribute()]
enum ModeFlags {
    A = 1
    B = 2
}
$mode = [ModeFlags]::A
$mode = $mode -bxor [ModeFlags]::B # add B
$hasB = $mode.HasFlag([ModeFlags]::B)
$mode = $mode -bxor [ModeFlags]::B # remove B
$hasBAfter = $mode.HasFlag([ModeFlags]::B)
if (-not $hasB -or $hasBAfter) {
    Write-Host "FAIL: Bitwise XOR flag toggle failed"
    exit 1
}
Write-Host "PASS"
exit 0
