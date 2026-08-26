# vybe-test: powershell/classes_this_chaining_and_scope/this_mutating_hashtable_member
class ConfigMap {
    [hashtable]$Map = @{}
    [void]SetVal([string]$k, [string]$v) { $this.Map[$k] = $v }
    [string]GetVal([string]$k) { return $this.Map[$k] }
}
$cm = [ConfigMap]::new()
$cm.SetVal("theme", "dark")
if ($cm.GetVal("theme") -ne "dark") {
    Write-Host "FAIL: ConfigMap `$this hashtable mutation failed"
    exit 1
}
Write-Host "PASS"
exit 0
