# vybe-test: powershell/enums_flags_attribute/enum_flags_tostring_comma_separated
[System.FlagsAttribute()]
enum ColorFlags {
    Red   = 1
    Green = 2
    Blue  = 4
}
$combo = [ColorFlags]::Red -bor [ColorFlags]::Blue
$str = $combo.ToString()
if ($str -ne "Red, Blue") {
    Write-Host "FAIL: Flags ToString expected 'Red, Blue', got '$str'"
    exit 1
}
Write-Host "PASS"
exit 0
