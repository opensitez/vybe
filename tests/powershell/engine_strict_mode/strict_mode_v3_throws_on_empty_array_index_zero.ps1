# vybe-test: powershell/engine_strict_mode/strict_mode_v3_throws_on_empty_array_index_zero
Set-StrictMode -Version 3.0

$emptyArray = @()

# Indexing element 0 on an empty array must throw IndexOutOfRangeException under v3.0
$caughtError = $null
try {
    $val = $emptyArray[0]
} catch {
    $caughtError = $_
}

if ($null -eq $caughtError) {
    Write-Host "FAIL: indexing element 0 on empty array did not throw under v3.0"
    exit 1
}

if (-not ($caughtError.Exception -is [System.IndexOutOfRangeException])) {
    Write-Host "FAIL: expected System.IndexOutOfRangeException, got: $($caughtError.Exception.GetType().FullName)"
    exit 1
}

Write-Host "PASS"
exit 0
