# vybe-test: powershell/enums_flags_attribute/enum_flags_parse_from_comma_string
[System.FlagsAttribute()]
enum FlagStyles {
    Bold   = 1
    Italic = 2
    Underline = 4
}
$parsed = [FlagStyles]"Bold, Underline"
if (-not $parsed.HasFlag([FlagStyles]::Bold) -or -not $parsed.HasFlag([FlagStyles]::Underline) -or $parsed.HasFlag([FlagStyles]::Italic)) {
    Write-Host "FAIL: Flags parse from comma string failed, got '$parsed'"
    exit 1
}
Write-Host "PASS"
exit 0
