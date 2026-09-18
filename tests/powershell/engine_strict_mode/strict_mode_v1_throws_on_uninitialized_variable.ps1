# vybe-test: powershell/engine_strict_mode/strict_mode_v1_throws_on_uninitialized_variable
Set-StrictMode -Version 1.0

$caughtError = $null
try {
    $val = $uninitializedTargetVar
} catch {
    $caughtError = $_
}

if ($null -eq $caughtError) {
    Write-Host "FAIL: Set-StrictMode -Version 1.0 did not throw on uninitialized variable access"
    exit 1
}

$isExpectedError = ($caughtError.FullyQualifiedErrorId -match "VariableIsUndefined") -or 
                    ($caughtError.Exception -is [System.Management.Automation.RuntimeException])

if (-not $isExpectedError) {
    Write-Host "FAIL: expected VariableIsUndefined error, got exception type: $($caughtError.Exception.GetType().FullName), error id: $($caughtError.FullyQualifiedErrorId)"
    exit 1
}

if (-not ($caughtError.Exception.Message -match "uninitializedTargetVar")) {
    Write-Host "FAIL: error message did not mention variable name: $($caughtError.Exception.Message)"
    exit 1
}

Write-Host "PASS"
exit 0
