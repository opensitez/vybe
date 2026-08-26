# vybe-test: powershell/classes_interface_implementation/implement_icomparable_compareto_method
class PriorityJob : System.IComparable {
    [int]$Priority
    PriorityJob([int]$p) { $this.Priority = $p }
    [int]CompareTo([object]$obj) {
        $other = [PriorityJob]$obj
        return $this.Priority.CompareTo($other.Priority)
    }
}
$j1 = [PriorityJob]::new(10)
$j2 = [PriorityJob]::new(20)
$cmp = $j1.CompareTo($j2)
if ($cmp -ge 0) {
    Write-Host "FAIL: IComparable CompareTo failed"
    exit 1
}
Write-Host "PASS"
exit 0
