# vybe-test: powershell/classes_base_method_calls/base_method_with_arguments
class BaseCalc {
    [int]Add([int]$a, [int]$b) { return $a + $b }
}
class SubCalc : BaseCalc {
    [int]AddThree([int]$a, [int]$b, [int]$c) {
        $twoSum = ([BaseCalc]$this).Add($a, $b)
        return ([BaseCalc]$this).Add($twoSum, $c)
    }
}
$sc = [SubCalc]::new()
$res = $sc.AddThree(10, 20, 30)
if ($res -ne 60) {
    Write-Host "FAIL: Base method with arguments failed, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
