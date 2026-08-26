# vybe-test: powershell/classes_custom_methods_overloading/overload_no_parameters_vs_parameters
class Stamp {
    [string]GetStamp() { return $this.GetStamp("default") }
    [string]GetStamp([string]$prefix) { return "$prefix-stamp" }
}
$st = [Stamp]::new()
if ($st.GetStamp() -ne "default-stamp" -or $st.GetStamp("custom") -ne "custom-stamp") {
    Write-Host "FAIL: Parameterless vs parameterized overload failed"
    exit 1
}
Write-Host "PASS"
exit 0
