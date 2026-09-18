# vybe-test: powershell/split_path_cmdlet/split_path_qualifier_missing_throws
# When -Qualifier is called on a path that does not contain a drive qualifier,
# Split-Path throws an error indicating that the path lacks a qualifier.
$threwExpected = $false

try {
    Split-Path "/usr/local/bin" -Qualifier -ErrorAction Stop
} catch {
    $threwExpected = $true
}

if (-not $threwExpected) {
    Write-Host "FAIL: Split-Path -Qualifier on path without qualifier did not throw"
    exit 1
}

Write-Host "PASS"
exit 0
