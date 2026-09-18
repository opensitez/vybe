# vybe-test: powershell/join_string_cmdlet/join_string_single_item_without_separator
$single = @('lonelyElement')

# A single item must NOT have the separator appended or prepended
$res = $single | Join-String -Separator '---SEPARATOR---'

if ($res -ne "lonelyElement") {
    Write-Host "FAIL: expected 'lonelyElement', got: '$res'"
    exit 1
}

Write-Host "PASS"
exit 0
