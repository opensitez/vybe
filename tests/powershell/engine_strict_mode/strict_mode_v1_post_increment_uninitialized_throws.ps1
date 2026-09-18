# vybe-test: powershell/engine_strict_mode/strict_mode_v1_post_increment_uninitialized_throws
Set-StrictMode -Version 1.0

$caughtError = $null
try {
    # In strict mode v1+, post-incrementing an uninitialized variable cannot read it to increment
    $uninitializedCounter++
} catch {
    $caughtError = $_
}

if ($null -eq $caughtError) {
    Write-Host "FAIL: post-incrementing uninitialized variable did not throw under v1.0"
    exit 1
}

if (-not ($caughtError.FullyQualifiedErrorId -match "VariableIsUndefined" -or 
          $caughtError.Exception -is [System.Management.Automation.RuntimeException])) {
    Write-Host "FAIL: expected VariableIsUndefined error, got: $($caughtError.FullyQualifiedErrorId)"
    exit 1
}

Write-Host "PASS"
exit 0
