# vybe-test: powershell/engine_strict_mode/strict_mode_v1_string_interpolation_uninitialized_evaluates_empty
Set-StrictMode -Version 1.0

# PowerShell Language Specification: string expandable tokens interpolate uninitialized variables as empty strings without throwing
$caught = $false
$result = $null
try {
    $result = "Start_${uninitializedForInterpolation}_End"
} catch {
    $caught = $true
}

if ($caught) {
    Write-Host "FAIL: string interpolation threw on uninitialized variable under v1.0"
    exit 1
}

if ($result -ne "Start__End") {
    Write-Host "FAIL: expected 'Start__End', got '$result'"
    exit 1
}

Write-Host "PASS"
exit 0
