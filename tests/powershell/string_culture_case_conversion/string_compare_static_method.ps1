# vybe-test: powershell/string_culture_case_conversion/string_compare_static_method
$res = [string]::Compare("ABC", "abc", $true) # ignoreCase = true
if ($res -ne 0) {
    Write-Host "FAIL: [string]::Compare with ignoreCase expected 0, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
