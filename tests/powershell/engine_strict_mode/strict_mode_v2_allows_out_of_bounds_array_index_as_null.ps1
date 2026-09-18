# vybe-test: powershell/engine_strict_mode/strict_mode_v2_allows_out_of_bounds_array_index_as_null
Set-StrictMode -Version 2.0

$arr = @("alpha", "beta", "gamma")

$caught = $false
try {
    # In Version 2.0, out-of-bounds indexing on collections is still permitted (deferred to v3.0)
    $val = $arr[99]
} catch {
    $caught = $true
}

if ($caught) {
    Write-Host "FAIL: Set-StrictMode -Version 2.0 should not throw on out-of-bounds array index"
    exit 1
}

if ($null -ne $val) {
    Write-Host "FAIL: expected out-of-bounds index under v2.0 to be `$null, got: $val"
    exit 1
}

# Verify valid array indexing still functions perfectly
if ($arr[0] -ne "alpha" -or $arr[2] -ne "gamma") {
    Write-Host "FAIL: valid array indexing failed under v2.0"
    exit 1
}

Write-Host "PASS"
exit 0
