# vybe-test: powershell/classes_interface_implementation/ienumerable_generic_getenumerator
class SimpleContainer : System.Collections.IEnumerable {
    [int[]]$Items = @(10, 20, 30)
    [System.Collections.IEnumerator]GetEnumerator() {
        return $this.Items.GetEnumerator()
    }
}
$sc = [SimpleContainer]::new()
$sum = 0
foreach ($item in $sc) { $sum += $item }
if ($sum -ne 60) {
    Write-Host "FAIL: IEnumerable GetEnumerator iteration failed, got $sum"
    exit 1
}
Write-Host "PASS"
exit 0
