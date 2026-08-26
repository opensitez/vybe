# vybe-test: powershell/error_record_exception_chaining/chained_exception_tostring_contains_inner_details
$inner = [System.FormatException]::new("BadFormatDetails")
$outer = [System.Exception]::new("OuterDetails", $inner)
$str = $outer.ToString()
if (-not ($str.Contains("BadFormatDetails") -and $str.Contains("OuterDetails"))) {
    Write-Host "FAIL: Chained exception ToString should mention both inner and outer messages"
    exit 1
}
Write-Host "PASS"
exit 0
