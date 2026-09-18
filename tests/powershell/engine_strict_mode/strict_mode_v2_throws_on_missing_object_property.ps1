# vybe-test: powershell/engine_strict_mode/strict_mode_v2_throws_on_missing_object_property
Set-StrictMode -Version 2.0

$obj = [pscustomobject]@{ ValidProp = 42 }

$caughtError = $null
try {
    # Version 2.0 explicitly forbids accessing non-existent properties
    $val = $obj.ThisPropertyDoesNotExist
} catch {
    $caughtError = $_
}

if ($null -eq $caughtError) {
    Write-Host "FAIL: Set-StrictMode -Version 2.0 did not throw on missing object property access"
    exit 1
}

$isPropertyError = ($caughtError.FullyQualifiedErrorId -match "PropertyNotFoundStrict") -or 
                   ($caughtError.Exception -is [System.Management.Automation.PropertyNotFoundException])

if (-not $isPropertyError) {
    Write-Host "FAIL: expected PropertyNotFoundStrict, got: $($caughtError.FullyQualifiedErrorId)"
    exit 1
}

# Verify valid property still resolves correctly
if ($obj.ValidProp -ne 42) {
    Write-Host "FAIL: valid property access corrupted under v2.0"
    exit 1
}

Write-Host "PASS"
exit 0
