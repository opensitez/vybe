# vybe-test: powershell/dynamic_property_lookup_by_variable/dynamic_property_case_insensitivity
$obj = [pscustomobject]@{ UniqueCode = "XYZ" }
$propUpper = "UNIQUECODE"
$propLower = "uniquecode"
if ($obj.$propUpper -ne "XYZ" -or $obj.$propLower -ne "XYZ") {
    Write-Host "FAIL: Dynamic property case-insensitivity failed"
    exit 1
}
Write-Host "PASS"
exit 0
