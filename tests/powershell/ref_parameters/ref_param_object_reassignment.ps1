# vybe-test: powershell/ref_parameters/ref_param_object_reassignment
function Reset-Obj([ref]$o) {
    $o.Value = [pscustomobject]@{ Status = "Reset" }
}
$obj = [pscustomobject]@{ Status = "Original" }
Reset-Obj ([ref]$obj)
if ($obj.Status -ne "Reset") {
    Write-Host "FAIL: object reference reassignment expected Status=Reset, got $($obj.Status)"
    exit 1
}
Write-Host "PASS"
exit 0
