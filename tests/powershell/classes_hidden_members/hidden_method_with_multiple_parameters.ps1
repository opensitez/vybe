# vybe-test: powershell/classes_hidden_members/hidden_method_with_multiple_parameters
class Combiner {
    hidden [string]JoinParts([string]$a, [string]$b, [string]$c) {
        return "$a-$b-$c"
    }
    [string]Run() { return $this.JoinParts("1", "2", "3") }
}
$c = [Combiner]::new()
if ($c.Run() -ne "1-2-3") {
    Write-Host "FAIL: Hidden method with multiple parameters failed"
    exit 1
}
Write-Host "PASS"
exit 0
