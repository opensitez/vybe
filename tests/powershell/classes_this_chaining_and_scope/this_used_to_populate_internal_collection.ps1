# vybe-test: powershell/classes_this_chaining_and_scope/this_used_to_populate_internal_collection
class RegistryClass {
    [System.Collections.Generic.List[string]]$Logs = [System.Collections.Generic.List[string]]::new()
    [void]LogSelf() {
        $this.Logs.Add("Logged at $([datetime]::UtcNow.Year)")
    }
}
$r = [RegistryClass]::new()
$r.LogSelf()
if ($r.Logs.Count -ne 1) {
    Write-Host "FAIL: `$this internal collection population failed"
    exit 1
}
Write-Host "PASS"
exit 0
