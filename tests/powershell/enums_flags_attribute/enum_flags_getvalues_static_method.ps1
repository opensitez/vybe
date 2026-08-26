# vybe-test: powershell/enums_flags_attribute/enum_flags_getvalues_static_method
[System.FlagsAttribute()]
enum ValFlags {
    A = 1
    B = 2
    C = 4
}
$vals = @([System.Enum]::GetValues([ValFlags]))
if ($vals.Length -ne 3 -or $vals[0] -ne [ValFlags]::A -or $vals[2] -ne [ValFlags]::C) {
    Write-Host "FAIL: Enum.GetValues failed"
    exit 1
}
Write-Host "PASS"
exit 0
