# vybe-test: powershell/convert_path_cmdlet/convert_path_dot_slash_prefix
# Prefixing with './' produces the identical resolved canonical path as without prefix
$withPrefix = Convert-Path "./tests"
$withoutPrefix = Convert-Path "tests"

if ($withPrefix -ne $withoutPrefix) {
    Write-Host "FAIL: './tests' ($withPrefix) did not match 'tests' ($withoutPrefix)"
    exit 1
}

Write-Host "PASS"
exit 0
