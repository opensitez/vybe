# vybe-test: powershell/type_accelerators/type_accelerator_regex
$re = [regex]"^(\d{3})-(\d{4})$"
$match = $re.Match("555-1234")
if (-not $match.Success) {
    Write-Host "FAIL: regex match failed"
    exit 1
}
if ($match.Groups[1].Value -ne "555") {
    Write-Host "FAIL: group 1 expected 555, got $($match.Groups[1].Value)"
    exit 1
}
if ($match.Groups[2].Value -ne "1234") {
    Write-Host "FAIL: group 2 expected 1234, got $($match.Groups[2].Value)"
    exit 1
}
Write-Host "PASS"
exit 0
