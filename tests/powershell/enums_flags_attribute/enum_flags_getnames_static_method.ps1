# vybe-test: powershell/enums_flags_attribute/enum_flags_getnames_static_method
[System.FlagsAttribute()]
enum NameFlags {
    Read = 1
    Write = 2
}
$names = @([System.Enum]::GetNames([NameFlags]))
if ($names.Length -ne 2 -or $names[0] -ne "Read" -or $names[1] -ne "Write") {
    Write-Host "FAIL: Enum.GetNames failed"
    exit 1
}
Write-Host "PASS"
exit 0
