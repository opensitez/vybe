# vybe-test: powershell/classes_polymorphism_and_casting/polymorphic_comparison_with_base_variable
class NumBase { [int]$N; NumBase([int]$n) { $this.N = $n } }
class SubNum : NumBase { SubNum([int]$n) : base($n) {} }
[NumBase]$a = [SubNum]::new(5)
[NumBase]$b = [SubNum]::new(5)
if ($a.N -ne $b.N) {
    Write-Host "FAIL: Polymorphic comparison failed"
    exit 1
}
Write-Host "PASS"
exit 0
