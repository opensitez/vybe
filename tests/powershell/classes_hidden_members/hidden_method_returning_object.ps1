# vybe-test: powershell/classes_hidden_members/hidden_method_returning_object
class ObjProducer {
    hidden [pscustomobject]CreateRaw() { return [pscustomobject]@{ Status = "Active" } }
    [string]GetStatus() { return $this.CreateRaw().Status }
}
$op = [ObjProducer]::new()
if ($op.GetStatus() -ne "Active") {
    Write-Host "FAIL: Hidden method returning object failed"
    exit 1
}
Write-Host "PASS"
exit 0
