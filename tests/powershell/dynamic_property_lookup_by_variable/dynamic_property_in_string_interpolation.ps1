# vybe-test: powershell/dynamic_property_lookup_by_variable/dynamic_property_in_string_interpolation
$p = "Title"
$obj = [pscustomobject]@{ Title = "Report" }
$str = "Document: $($obj.$p)"
if ($str -ne "Document: Report") {
    Write-Host "FAIL: Dynamic property in string interpolation failed"
    exit 1
}
Write-Host "PASS"
exit 0
