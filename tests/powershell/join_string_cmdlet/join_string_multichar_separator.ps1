# vybe-test: powershell/join_string_cmdlet/join_string_multichar_separator
$tokens = @('TokenA', 'TokenB', 'TokenC')

# -Separator can be a multi-character delimiter string
$res = $tokens | Join-String -Separator ' <=> '

if ($res -ne "TokenA <=> TokenB <=> TokenC") {
    Write-Host "FAIL: expected 'TokenA <=> TokenB <=> TokenC', got: '$res'"
    exit 1
}

Write-Host "PASS"
exit 0
