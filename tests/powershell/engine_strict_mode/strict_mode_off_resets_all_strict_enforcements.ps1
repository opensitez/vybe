# vybe-test: powershell/engine_strict_mode/strict_mode_off_resets_all_strict_enforcements
# Enable strictest mode first
Set-StrictMode -Version 3.0

# Now explicitly disable strict mode
Set-StrictMode -Off

# 1. Uninitialized variables must evaluate to $null without throwing
$uninitializedVal = $null
$caughtVar = $false
try {
    $uninitializedVal = $brandNewUninitializedVar
} catch {
    $caughtVar = $true
}

if ($caughtVar -or $null -ne $uninitializedVal) {
    Write-Host "FAIL: Set-StrictMode -Off did not restore permissive uninitialized variable behavior"
    exit 1
}

# 2. Missing object properties must evaluate to $null without throwing
$testObj = [pscustomobject]@{ Known = "val" }
$missingPropVal = $null
$caughtProp = $false
try {
    $missingPropVal = $testObj.NonExistentProp
} catch {
    $caughtProp = $true
}

if ($caughtProp -or $null -ne $missingPropVal) {
    Write-Host "FAIL: Set-StrictMode -Off did not restore permissive missing property behavior"
    exit 1
}

# 3. Out-of-bounds array access must evaluate to $null without throwing
$testArr = @(1, 2)
$missingArrVal = $null
$caughtArr = $false
try {
    $missingArrVal = $testArr[100]
} catch {
    $caughtArr = $true
}

if ($caughtArr -or $null -ne $missingArrVal) {
    Write-Host "FAIL: Set-StrictMode -Off did not restore permissive array indexing behavior"
    exit 1
}

Write-Host "PASS"
exit 0
