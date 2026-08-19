# vybe-test: powershell/string_literal_modes/literal_plus_interpolation_split
$section = 'core'
$result = "prefix_${section}_suffix"
if ($result -ne 'prefix_core_suffix') {
    Write-Host "FAIL: interpolation split around literal text failed: '$result'"
    exit 1
}

Write-Host 'PASS'
exit 0
