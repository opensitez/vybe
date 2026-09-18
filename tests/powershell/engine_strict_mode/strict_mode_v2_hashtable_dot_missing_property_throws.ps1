# vybe-test: powershell/engine_strict_mode/strict_mode_v2_hashtable_dot_missing_property_throws
Set-StrictMode -Version 2.0

$hash = @{ ExistingKey = "present" }

# Accessing an existing key via dot syntax succeeds
if ($hash.ExistingKey -ne "present") {
    Write-Host "FAIL: existing key dot access failed under v2.0"
    exit 1
}

# Accessing a missing key on a hashtable via dot syntax must throw under v2.0
$caughtError = $null
try {
    $val = $hash.MissingKey
} catch {
    $caughtError = $_
}

if ($null -eq $caughtError) {
    Write-Host "FAIL: hashtable dot access for missing key did not throw under v2.0"
    exit 1
}

if (-not ($caughtError.FullyQualifiedErrorId -match "PropertyNotFoundStrict" -or 
          $caughtError.Exception -is [System.Management.Automation.PropertyNotFoundException])) {
    Write-Host "FAIL: expected PropertyNotFoundStrict, got: $($caughtError.FullyQualifiedErrorId)"
    exit 1
}

Write-Host "PASS"
exit 0
