# vybe-test: powershell/classes_hidden_members/hidden_hashtable_property
class HiddenPropClass {
    hidden [hashtable]$Config
    HiddenPropClass([hashtable]$c) { $this.Config = $c }
    [string]GetConfigKey([string]$k) { return $this.Config[$k] }
}
$inst = [HiddenPropClass]::new(@{ env = "prod" })
if ($inst.GetConfigKey("env") -ne "prod") {
    Write-Host "FAIL: Hidden hashtable property failed"
    exit 1
}
Write-Host "PASS"
exit 0
