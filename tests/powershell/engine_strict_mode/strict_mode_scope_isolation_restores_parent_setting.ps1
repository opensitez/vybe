# vybe-test: powershell/engine_strict_mode/strict_mode_scope_isolation_restores_parent_setting
Set-StrictMode -Off

function Invoke-StrictChildScope {
    Set-StrictMode -Version 3.0
    $childCaught = $null
    try {
        $x = $uninitializedInChild
    } catch {
        $childCaught = $_
    }
    if ($null -eq $childCaught) {
        Write-Host "FAIL: child scope with Set-StrictMode -Version 3.0 failed to throw on uninitialized variable"
        exit 1
    }
    if (-not ($childCaught.FullyQualifiedErrorId -match "VariableIsUndefined" -or 
              $childCaught.Exception -is [System.Management.Automation.RuntimeException])) {
        Write-Host "FAIL: child scope threw unexpected error: $($childCaught.FullyQualifiedErrorId)"
        exit 1
    }
}

Invoke-StrictChildScope

# Scope isolation guarantee: the parent scope must remain in Set-StrictMode -Off
$parentCaught = $false
try {
    $parentVal = $uninitializedInParentScope
} catch {
    $parentCaught = $true
}

if ($parentCaught) {
    Write-Host "FAIL: parent scope strict mode setting was overwritten by child function"
    exit 1
}

if ($null -ne $parentVal) {
    Write-Host "FAIL: expected parent scope uninitialized variable to be `$null, got: $parentVal"
    exit 1
}

Write-Host "PASS"
exit 0
