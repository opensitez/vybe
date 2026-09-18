# vybe-test: powershell/engine_strict_mode/strict_mode_v1_allows_initialized_null_variable
Set-StrictMode -Version 1.0

# A variable explicitly assigned $null is initialized and must not throw under v1.0
$explicitNull = $null

$caught = $false
try {
    $val = $explicitNull
} catch {
    $caught = $true
}

if ($caught) {
    Write-Host "FAIL: reading variable explicitly set to `$null threw under v1.0"
    exit 1
}

if ($null -ne $val) {
    Write-Host "FAIL: expected variable value to be `$null, got: $val"
    exit 1
}

Write-Host "PASS"
exit 0
