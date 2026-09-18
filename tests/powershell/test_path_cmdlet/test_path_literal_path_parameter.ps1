# vybe-test: powershell/test_path_cmdlet/test_path_literal_path_parameter
# -LiteralPath tests the exact path without wildcard interpretation
$res = Test-Path -LiteralPath "tests"

if ($res -ne $true) {
    Write-Host "FAIL: Test-Path -LiteralPath returned `$false for existing directory"
    exit 1
}

Write-Host "PASS"
exit 0
