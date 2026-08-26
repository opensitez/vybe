# vybe-test: powershell/enums_flags_attribute/enum_flags_cast_from_integer
[System.FlagsAttribute()]
enum BitFlagsEnum {
    One = 1
    Two = 2
    Four = 4
}
$f = [BitFlagsEnum]5 # 1 + 4 => One, Four
if ($f.ToString() -ne "One, Four") {
    Write-Host "FAIL: Integer cast to flags enum failed, got '$f'"
    exit 1
}
Write-Host "PASS"
exit 0
