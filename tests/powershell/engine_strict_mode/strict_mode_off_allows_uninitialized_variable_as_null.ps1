# vybe-test: powershell/engine_strict_mode/strict_mode_off_allows_uninitialized_variable_as_null
Set-StrictMode -Off

$caught = $false
try {
    $val = $uninitializedVariableThatDoesNotExist
} catch {
    $caught = $true
}

if ($caught) {
    Write-Host "FAIL: Set-StrictMode -Off threw when reading uninitialized variable"
    exit 1
}

if ($null -ne $val) {
    Write-Host "FAIL: expected uninitialized variable to evaluate to `$null under -Off, got: $val"
    exit 1
}

# In permissive mode, uninitialized variables interpolate as empty strings
$interpolated = "Prefix_${val}_Suffix"
if ($interpolated -ne "Prefix__Suffix") {
    Write-Host "FAIL: expected empty interpolation 'Prefix__Suffix', got '$interpolated'"
    exit 1
}

Write-Host "PASS"
exit 0
