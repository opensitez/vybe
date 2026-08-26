# vybe-test: powershell/classes_interface_implementation/icomparable_equality_zero_return
class EqItem : System.IComparable {
    [int]$Key
    EqItem([int]$k) { $this.Key = $k }
    [int]CompareTo([object]$obj) {
        $other = [EqItem]$obj
        return $this.Key.CompareTo($other.Key)
    }
}
$a = [EqItem]::new(5)
$b = [EqItem]::new(5)
if ($a.CompareTo($b) -ne 0) {
    Write-Host "FAIL: IComparable 0 return failed"
    exit 1
}
Write-Host "PASS"
exit 0
