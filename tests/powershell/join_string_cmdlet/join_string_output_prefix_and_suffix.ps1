# vybe-test: powershell/join_string_cmdlet/join_string_output_prefix_and_suffix
# -OutputPrefix prepends text and -OutputSuffix appends text to the joined string
$res = 1..3 | Join-String -Separator ', ' -OutputPrefix '[' -OutputSuffix ']'

if ($res -ne "[1, 2, 3]") {
    Write-Host "FAIL: expected '[1, 2, 3]', got: '$res'"
    exit 1
}

Write-Host "PASS"
exit 0
