# vybe-test: powershell/variable_expansion/expand_property_in_string
$obj = [pscustomobject]@{ Name = 'y' }
if ("Hello $($obj.Name)" -ne 'Hello y') {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
