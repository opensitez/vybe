# vybe-test: powershell/classes_interface_implementation/implement_multiple_interfaces
class MultiContract : System.IDisposable, System.IComparable {
    [int]$Score
    MultiContract([int]$s) { $this.Score = $s }
    [void]Dispose() {}
    [int]CompareTo([object]$obj) {
        $other = [MultiContract]$obj
        return $this.Score.CompareTo($other.Score)
    }
}
$mc = [MultiContract]::new(50)
if ($mc -isnot [System.IDisposable] -or $mc -isnot [System.IComparable]) {
    Write-Host "FAIL: Multiple interfaces conformance check failed"
    exit 1
}
Write-Host "PASS"
exit 0
