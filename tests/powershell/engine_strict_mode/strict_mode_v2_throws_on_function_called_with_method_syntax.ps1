# vybe-test: powershell/engine_strict_mode/strict_mode_v2_throws_on_function_called_with_method_syntax
Set-StrictMode -Version 2.0

function Multiply-Values($a, $b) {
    return $a * $b
}

# Standard invocation must work fine under v2.0
$normalResult = Multiply-Values 3 4
if ($normalResult -ne 12) {
    Write-Host "FAIL: standard function call failed under v2.0, expected 12, got $normalResult"
    exit 1
}

# In StrictMode v2.0+, calling a function with parentheses like Multiply-Values(3, 4) is an engine error
$caughtError = $null
try {
    Invoke-Command { Multiply-Values(3, 4) }
} catch {
    $caughtError = $_
}

if ($null -eq $caughtError) {
    Write-Host "FAIL: Set-StrictMode -Version 2.0 did not throw on function called with parentheses syntax"
    exit 1
}

if (-not ($caughtError.FullyQualifiedErrorId -match "StrictModeFunctionCallWithParens" -or 
          $caughtError.Exception.Message -match "parentheses")) {
    Write-Host "FAIL: expected StrictModeFunctionCallWithParens error, got: $($caughtError.FullyQualifiedErrorId)"
    exit 1
}

Write-Host "PASS"
exit 0
