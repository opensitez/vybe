# vybe-test: powershell/exceptions_throw_expression_vs_statement/throw_in_constructor_body
class StrictTargetConstructor {
    StrictTargetConstructor([int]$val) {
        if ($val -lt 0) { throw [System.ArgumentOutOfRangeException]::new("val must be positive") }
    }
}
$caught = $false
try {
    $x = [StrictTargetConstructor]::new(-5)
} catch [System.ArgumentOutOfRangeException] {
    $caught = $true
}
if (-not $caught) {
    Write-Host "FAIL: Throw in constructor body failed"
    exit 1
}
Write-Host "PASS"
exit 0
