# vybe-test: powershell/regex_named_capture_groups/groupnames_property_inspection
$re = [regex]::new("(?<first>[a-z]+)-(?<second>\d+)")
$names = @($re.GetGroupNames())
if (-not ($names -contains "first") -or -not ($names -contains "second") -or -not ($names -contains "0")) {
    Write-Host "FAIL: GetGroupNames failed"
    exit 1
}
Write-Host "PASS"
exit 0
