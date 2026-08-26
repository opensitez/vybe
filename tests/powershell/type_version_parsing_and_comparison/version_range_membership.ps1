# vybe-test: powershell/type_version_parsing_and_comparison/version_range_membership
$target = [version]"2.5.0"
$inRange = ($target -ge [version]"2.0.0") -and ($target -lt [version]"3.0.0")
if (-not $inRange) {
    Write-Host "FAIL: Version range membership test failed"
    exit 1
}
Write-Host "PASS"
exit 0
