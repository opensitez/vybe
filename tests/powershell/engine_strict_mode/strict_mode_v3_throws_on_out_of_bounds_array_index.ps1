# vybe-test: powershell/engine_strict_mode/strict_mode_v3_throws_on_out_of_bounds_array_index
Set-StrictMode -Version 3.0

$arr = @(100, 200, 300)

# Valid in-bound access must succeed
if ($arr[1] -ne 200) {
    Write-Host "FAIL: in-bounds array access failed under v3.0, expected 200, got $($arr[1])"
    exit 1
}

# Version 3.0 strictly forbids indexing out of bounds on arrays/collections
$caughtError = $null
try {
    $val = $arr[10]
} catch {
    $caughtError = $_
}

if ($null -eq $caughtError) {
    Write-Host "FAIL: Set-StrictMode -Version 3.0 did not throw on out-of-bounds array index"
    exit 1
}

if (-not ($caughtError.Exception -is [System.IndexOutOfRangeException])) {
    Write-Host "FAIL: expected System.IndexOutOfRangeException under v3.0, got: $($caughtError.Exception.GetType().FullName)"
    exit 1
}

Write-Host "PASS"
exit 0
