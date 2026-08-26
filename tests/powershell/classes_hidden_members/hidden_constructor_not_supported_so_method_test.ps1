# vybe-test: powershell/classes_hidden_members/hidden_constructor_not_supported_so_method_test
class HelperClass {
    hidden [void]Log([string]$msg) {}
    [void]DoWork() { $this.Log("working") }
}
$h = [HelperClass]::new()
$h.DoWork()
Write-Host "PASS"
exit 0
