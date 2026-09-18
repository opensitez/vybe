# vybe-test: powershell/engine_strict_mode/strict_mode_v1_permits_missing_object_property
Set-StrictMode -Version 1.0

# In Version 1.0, property checks are NOT enabled yet; only uninitialized variables are checked.
$obj = [pscustomobject]@{ ExistingProp = "hello"; Count = 5 }

$caught = $false
try {
    $propVal = $obj.NonExistentProperty
} catch {
    $caught = $true
}

if ($caught) {
    Write-Host "FAIL: Set-StrictMode -Version 1.0 should not throw on missing property access"
    exit 1
}

if ($null -ne $propVal) {
    Write-Host "FAIL: expected missing property to return `$null under v1.0, got: $propVal"
    exit 1
}

# Verify existing properties remain fully accessible under v1.0
if ($obj.ExistingProp -ne "hello" -or $obj.Count -ne 5) {
    Write-Host "FAIL: existing properties corrupted under v1.0"
    exit 1
}

Write-Host "PASS"
exit 0
