# vybe-test: powershell/pscustomobject_literals/pscustomobject_null_property
$obj = [pscustomobject]@{ NullProp = $null }
$prop = $obj.psobject.Properties["NullProp"]
if ($prop -eq $null) {
    Write-Host "FAIL: NullProp property missing from object metadata"
    exit 1
}
if ($prop.Value -ne $null) {
    Write-Host "FAIL: NullProp value expected null"
    exit 1
}
Write-Host "PASS"
exit 0
