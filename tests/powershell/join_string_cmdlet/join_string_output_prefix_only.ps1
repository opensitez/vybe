# vybe-test: powershell/join_string_cmdlet/join_string_output_prefix_only
# -OutputPrefix alone prepends without requiring -OutputSuffix
$res = 1..3 | Join-String -Separator '-' -OutputPrefix 'ID:'

if ($res -ne "ID:1-2-3") {
    Write-Host "FAIL: expected 'ID:1-2-3', got: '$res'"
    exit 1
}

Write-Host "PASS"
exit 0
