# vybe-test: powershell/language_null_coalescing_and_assignment/null_coalescing_in_class_method
class CoalesceClass {
    [string]$Cached = "InitValue"
    [string]GetOrSet([string]$defVal) {
        if ($this.Cached -eq $null) { $this.Cached = $defVal }
        return $this.Cached
    }
}
$cc = [CoalesceClass]::new()
$r1 = $cc.GetOrSet("InitValue")
if ($r1 -ne "InitValue") {
    Write-Host "FAIL: Coalesce in class method failed"
    exit 1
}
Write-Host "PASS"
exit 0
