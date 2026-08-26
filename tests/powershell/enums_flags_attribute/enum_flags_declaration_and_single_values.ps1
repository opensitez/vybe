# vybe-test: powershell/enums_flags_attribute/enum_flags_declaration_and_single_values
[System.FlagsAttribute()]
enum FileAccessMode {
    None    = 0
    Read    = 1
    Write   = 2
    Execute = 4
}
$r = [FileAccessMode]::Read
$w = [FileAccessMode]::Write
if ($r.value__ -ne 1 -or $w.value__ -ne 2) {
    Write-Host "FAIL: Flags enum single values failed"
    exit 1
}
Write-Host "PASS"
exit 0
