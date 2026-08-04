# vybe-test: powershell/strings/string_regex_replace_capture_group
$input = "John Smith"
$swapped = $input -replace "(\w+)\s+(\w+)", '$2, $1'
if ($swapped -ne "Smith, John") {
    Write-Host "FAIL: expected 'Smith, John', got '$swapped'"
    exit 1
}
Write-Host "PASS"
exit 0
