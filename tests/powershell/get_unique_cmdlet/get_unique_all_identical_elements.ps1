# vybe-test: powershell/get_unique_cmdlet/get_unique_all_identical_elements
# A collection consisting entirely of identical elements collapses to a single value
$items = @(42, 42, 42, 42)
$res = @($items | Get-Unique)

if ($res.Count -ne 1) {
    Write-Host "FAIL: expected 1 unique element, got $($res.Count)"
    exit 1
}

if ($res[0] -ne 42) {
    Write-Host "FAIL: expected value 42, got: $($res[0])"
    exit 1
}

Write-Host "PASS"
exit 0
