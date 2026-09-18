# vybe-test: powershell/engine_strict_mode/strict_mode_latest_enforces_version_three_rules
Set-StrictMode -Version Latest

# 'Latest' maps to the highest implemented version (currently v3.0 rules),
# which must enforce out-of-bounds array checks
$arr = @(1, 2)

$caughtError = $null
try {
    $val = $arr[10]
} catch {
    $caughtError = $_
}

if ($null -eq $caughtError) {
    Write-Host "FAIL: Set-StrictMode -Version Latest did not throw on out-of-bounds array access"
    exit 1
}

if (-not ($caughtError.Exception -is [System.IndexOutOfRangeException])) {
    Write-Host "FAIL: expected System.IndexOutOfRangeException under Latest, got: $($caughtError.Exception.GetType().FullName)"
    exit 1
}

Write-Host "PASS"
exit 0
