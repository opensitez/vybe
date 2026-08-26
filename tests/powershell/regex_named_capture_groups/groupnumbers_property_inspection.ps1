# vybe-test: powershell/regex_named_capture_groups/groupnumbers_property_inspection
$re = [regex]::new("(?<alpha>[A-Z]+)(?<num>\d+)")
$num1 = $re.GroupNumberFromName("alpha")
$num2 = $re.GroupNumberFromName("num")
if ($num1 -ne 1 -or $num2 -ne 2) {
    Write-Host "FAIL: GroupNumberFromName failed"
    exit 1
}
Write-Host "PASS"
exit 0
