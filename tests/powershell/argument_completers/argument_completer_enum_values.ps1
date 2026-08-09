# vybe-test: powershell/argument_completers/argument_completer_enum_values
$completer = {
    param($enumType)
    [enum]::GetNames($enumType)
}
$names = @(&$completer ([System.DayOfWeek]))
if ($names.Count -ne 7 -or $names[0] -ne "Sunday") {
    Write-Host "FAIL: enum argument completer expected 7 day names, starting with Sunday"
    exit 1
}
Write-Host "PASS"
exit 0
