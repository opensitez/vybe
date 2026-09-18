# vybe-test: powershell/classes/class_static_method_invocation
# Static methods on PowerShell classes are invoked using the [ClassName]::MethodName() syntax
class CalculationUtilities {
    static [int] ComputeHypotenuseSquared([int]$a, [int]$b) {
        return ($a * $a) + ($b * $b)
    }
}

$result = [CalculationUtilities]::ComputeHypotenuseSquared(3, 4)

if ($result -ne 25) {
    Write-Host "FAIL: expected 25, got $result"
    exit 1
}

Write-Host "PASS"
exit 0
