# vybe-test: powershell/null_conditional/null_conditional_property_null
$obj = $null
$res = ${obj}?.Name
if ($res -ne $null) {
    Write-Host "FAIL: null conditional property expected null, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
