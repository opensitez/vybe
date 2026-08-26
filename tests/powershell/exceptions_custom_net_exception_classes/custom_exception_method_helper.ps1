# vybe-test: powershell/exceptions_custom_net_exception_classes/custom_exception_method_helper
class CalcException : System.Exception {
    [int]$OperandA; [int]$OperandB
    CalcException([int]$a, [int]$b, [string]$m) : base($m) {
        $this.OperandA = $a
        $this.OperandB = $b
    }
    [string]GetDiagnostic() {
        return "Failed with A=$($this.OperandA), B=$($this.OperandB)"
    }
}
$ce = [CalcException]::new(10, 0, "Divide error")
if ($ce.GetDiagnostic() -ne "Failed with A=10, B=0") {
    Write-Host "FAIL: Custom exception method helper failed"
    exit 1
}
Write-Host "PASS"
exit 0
