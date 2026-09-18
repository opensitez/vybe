# vybe-test: powershell/engine_strict_mode/strict_mode_v3_throws_on_negative_index_out_of_bounds
Set-StrictMode -Version 3.0

$arr = @("first", "second")

# Valid negative index [-1] resolves correctly
if ($arr[-1] -ne "second") {
    Write-Host "FAIL: valid negative indexing [-1] failed under v3.0"
    exit 1
}

# Negative index that exceeds collection length (-3 on a 2-element array) must throw IndexOutOfRangeException
$caughtError = $null
try {
    $val = $arr[-3]
} catch {
    $caughtError = $_
}

if ($null -eq $caughtError) {
    Write-Host "FAIL: out-of-bounds negative index did not throw under v3.0"
    exit 1
}

if (-not ($caughtError.Exception -is [System.IndexOutOfRangeException])) {
    Write-Host "FAIL: expected System.IndexOutOfRangeException, got: $($caughtError.Exception.GetType().FullName)"
    exit 1
}

Write-Host "PASS"
exit 0
