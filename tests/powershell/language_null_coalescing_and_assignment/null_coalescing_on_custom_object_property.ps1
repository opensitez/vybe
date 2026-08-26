# vybe-test: powershell/language_null_coalescing_and_assignment/null_coalescing_on_custom_object_property
$obj = [pscustomobject]@{ SetProp = "Hello"; NullProp = $null }
$r1 = $obj.NullProp ?? "DefaultProp"
$r2 = $obj.SetProp ?? "DefaultProp"
if ($r1 -ne "DefaultProp" -or $r2 -ne "Hello") {
    Write-Host "FAIL: ?? on custom object property failed, r1='$r1', r2='$r2'"
    exit 1
}
Write-Host "PASS"
exit 0
