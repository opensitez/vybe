# vybe-test: powershell/classes_interface_implementation/sorting_objects_implementing_icomparable
class NumberItem : System.IComparable {
    [int]$Val
    NumberItem([int]$v) { $this.Val = $v }
    [int]CompareTo([object]$obj) {
        $other = [NumberItem]$obj
        return $this.Val.CompareTo($other.Val)
    }
}
$items = @([NumberItem]::new(30), [NumberItem]::new(10), [NumberItem]::new(20))
[System.Array]::Sort($items)
if ($items[0].Val -ne 10 -or $items[1].Val -ne 20 -or $items[2].Val -ne 30) {
    Write-Host "FAIL: [Array]::Sort on IComparable objects failed"
    exit 1
}
Write-Host "PASS"
exit 0
