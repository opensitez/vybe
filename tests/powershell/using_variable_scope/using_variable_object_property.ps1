# vybe-test: powershell/using_variable_scope/using_variable_object_property
$obj = [pscustomobject]@{ Setting = "Active" }
$sb = { ($using:obj).Setting }
$res = &$sb
if ($res -ne "Active") {
    Write-Host "FAIL: object property (\$using:obj).Setting expected 'Active', got '$res'"
    exit 1
}
Write-Host "PASS"
exit 0
