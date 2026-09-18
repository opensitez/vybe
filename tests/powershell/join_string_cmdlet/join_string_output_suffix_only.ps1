# vybe-test: powershell/join_string_cmdlet/join_string_output_suffix_only
# -OutputSuffix alone appends without requiring -OutputPrefix
$res = 1..3 | Join-String -Separator '-' -OutputSuffix ';'

if ($res -ne "1-2-3;") {
    Write-Host "FAIL: expected '1-2-3;', got: '$res'"
    exit 1
}

Write-Host "PASS"
exit 0
