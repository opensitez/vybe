# vybe-test: powershell/subexpressions/subexpression_in_property
$obj = [pscustomobject]@{ Value = $(1 + 1) }
if ($obj.Value -ne 2) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
