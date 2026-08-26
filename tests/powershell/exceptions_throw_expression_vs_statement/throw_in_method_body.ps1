# vybe-test: powershell/exceptions_throw_expression_vs_statement/throw_in_method_body
class ThrowerMethodClass {
    [int]Divide([int]$a, [int]$b) {
        if ($b -eq 0) { throw [System.DivideByZeroException]::new("b cannot be zero") }
        return [int]($a / $b)
    }
}
$tm = [ThrowerMethodClass]::new()
$r1 = $tm.Divide(10, 2)
$caught = $false
try {
    $r2 = $tm.Divide(10, 0)
} catch [System.DivideByZeroException] {
    $caught = $true
}
if ($r1 -ne 5 -or -not $caught) {
    Write-Host "FAIL: Throw in method body failed"
    exit 1
}
Write-Host "PASS"
exit 0
