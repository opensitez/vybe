# vybe-test: powershell/join_string_cmdlet/join_string_empty_string_elements_included
$items = @('alpha', '', 'omega')

# While $null items are skipped, empty strings [string]"" are NOT skipped and produce consecutive separators
$res = $items | Join-String -Separator ','

if ($res -ne "alpha,,omega") {
    Write-Host "FAIL: empty string element was unexpectedly skipped, expected 'alpha,,omega', got: '$res'"
    exit 1
}

Write-Host "PASS"
exit 0
