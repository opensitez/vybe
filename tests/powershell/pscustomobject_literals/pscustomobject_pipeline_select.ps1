# vybe-test: powershell/pscustomobject_literals/pscustomobject_pipeline_select
$obj = [pscustomobject]@{ Alpha = "A"; Beta = "B" }
$sub = $obj | Select-Object Alpha
if ($sub.Alpha -ne "A" -or $sub.psobject.Properties["Beta"] -ne $null) {
    Write-Host "FAIL: Select-Object expected only Alpha property remaining"
    exit 1
}
Write-Host "PASS"
exit 0
