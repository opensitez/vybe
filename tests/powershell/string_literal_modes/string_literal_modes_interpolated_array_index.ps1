# vybe-test: powershell/string_literal_modes/interpolated_array_index
$items = @('first', 'second', 'third')
$result = "${items[1]}"
if ($result -ne 'second') {
    Write-Host "FAIL: expected second, got '$result'"
    exit 1
}

Write-Host 'PASS'
exit 0
