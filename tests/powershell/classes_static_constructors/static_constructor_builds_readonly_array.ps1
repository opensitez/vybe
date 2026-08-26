# vybe-test: powershell/classes_static_constructors/static_constructor_builds_readonly_array
class ReadOnlyNumbers {
    static [int[]]$Primes
    static ReadOnlyNumbers() {
        [ReadOnlyNumbers]::Primes = @(2, 3, 5, 7, 11, 13)
    }
}
$p = [ReadOnlyNumbers]::Primes
if ($p.Length -ne 6 -or $p[0] -ne 2 -or $p[5] -ne 13) {
    Write-Host "FAIL: Static primes array failed"
    exit 1
}
Write-Host "PASS"
exit 0
