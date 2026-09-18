# vybe-test: powershell/convert_path_cmdlet/convert_path_double_slash_normalization
# Redundant duplicate slashes in a relative path are normalized into a canonical path
$res = Convert-Path "tests//powershell"
$expected = Convert-Path "tests/powershell"

if ($res -ne $expected) {
    Write-Host "FAIL: double slashes not normalized, expected '$expected', got: '$res'"
    exit 1
}

Write-Host "PASS"
exit 0
