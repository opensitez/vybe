# vybe-test: powershell/enums_flags_attribute/enum_flags_tryparse_method
[System.FlagsAttribute()]
enum PermLevels { None = 0; Read = 1; Write = 2; Execute = 4 }
$parsed = [PermLevels]::None
$ok = [System.Enum]::TryParse([PermLevels], "Read, Write", [ref]$parsed)
if (-not $ok -or $parsed -ne ([PermLevels]::Read -bor [PermLevels]::Write)) {
    Write-Host "FAIL: Enum.TryParse with flags failed"
    exit 1
}
Write-Host "PASS"
exit 0
