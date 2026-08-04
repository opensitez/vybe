# vybe-test: powershell/classes/class_interface_icomparable
class Temperature : System.IComparable {
    [double]$Celsius
    Temperature([double]$c) { $this.Celsius = $c }
    [int]CompareTo([object]$other) {
        return $this.Celsius.CompareTo($other.Celsius)
    }
}
$cold = [Temperature]::new(0)
$hot  = [Temperature]::new(100)
if ($cold.CompareTo($hot) -ge 0) { Write-Host "FAIL: cold should be less than hot"; exit 1 }
if ($hot.CompareTo($cold) -le 0) { Write-Host "FAIL: hot should be greater than cold"; exit 1 }
if ($cold.CompareTo($cold) -ne 0) { Write-Host "FAIL: equal should return 0"; exit 1 }
Write-Host "PASS"
exit 0
