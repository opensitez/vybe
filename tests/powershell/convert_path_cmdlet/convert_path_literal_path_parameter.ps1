# vybe-test: powershell/convert_path_cmdlet/convert_path_literal_path_parameter
# -LiteralPath resolves exact path without interpreting wildcard characters
$res = Convert-Path -LiteralPath "tests"

if ($null -eq $res) {
    Write-Host "FAIL: Convert-Path -LiteralPath returned `$null"
    exit 1
}

if ($res -notmatch "tests$") {
    Write-Host "FAIL: expected resolved path ending in 'tests', got: '$res'"
    exit 1
}

Write-Host "PASS"
exit 0
