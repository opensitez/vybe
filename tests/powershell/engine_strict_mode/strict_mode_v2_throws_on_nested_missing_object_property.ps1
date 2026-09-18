# vybe-test: powershell/engine_strict_mode/strict_mode_v2_throws_on_nested_missing_object_property
Set-StrictMode -Version 2.0

$nested = [pscustomobject]@{
    Child = [pscustomobject]@{
        Leaf = "existingLeaf"
    }
}

# The parent and intermediate child exist
if ($nested.Child.Leaf -ne "existingLeaf") {
    Write-Host "FAIL: valid nested property resolution failed under v2.0"
    exit 1
}

# Attempting to read a missing property on the child object must throw PropertyNotFoundStrict
$caughtError = $null
try {
    $val = $nested.Child.MissingLeafProperty
} catch {
    $caughtError = $_
}

if ($null -eq $caughtError) {
    Write-Host "FAIL: accessing missing nested property did not throw under v2.0"
    exit 1
}

if (-not ($caughtError.FullyQualifiedErrorId -match "PropertyNotFoundStrict" -or 
          $caughtError.Exception -is [System.Management.Automation.PropertyNotFoundException])) {
    Write-Host "FAIL: expected PropertyNotFoundStrict, got: $($caughtError.FullyQualifiedErrorId)"
    exit 1
}

Write-Host "PASS"
exit 0
