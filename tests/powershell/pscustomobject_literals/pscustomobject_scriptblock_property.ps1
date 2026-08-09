# vybe-test: powershell/pscustomobject_literals/pscustomobject_scriptblock_property
$obj = [pscustomobject]@{ Code = { param($x) $x + 1 } }
$val = &($obj.Code) 5
if ($val -ne 6) {
    Write-Host "FAIL: scriptblock property execution expected 6, got $val"
    exit 1
}
Write-Host "PASS"
exit 0
