# vybe-test: powershell/string_literal_modes/nested_expansion_failure
$name = 'value'
$result = '$($name)'
if ($result -ne '$($name)') {
    Write-Host "FAIL: nested expansion text expected as literal: '$result'"
    exit 1
}

Write-Host 'PASS'
exit 0
