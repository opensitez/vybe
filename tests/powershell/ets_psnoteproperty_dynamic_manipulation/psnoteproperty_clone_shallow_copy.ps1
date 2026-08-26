# vybe-test: powershell/ets_psnoteproperty_dynamic_manipulation/psnoteproperty_clone_shallow_copy
$orig = [pscustomobject]@{ Name = "Original" }
$copy = $orig.PSObject.Copy()
$copy.Name = "Modified"
if ($orig.Name -ne "Original" -or $copy.Name -ne "Modified") {
    Write-Host "FAIL: PSObject.Copy() shallow copy note separation failed"
    exit 1
}
Write-Host "PASS"
exit 0
