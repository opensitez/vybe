# vybe-test: powershell/regex/regex_replace_pattern
$str = "the cat sat on the mat"
$result = $str -replace "[cm]at", "dog"
if ($result -ne "the dog sat on the dog") {
    Write-Host "FAIL: got '$result'"
    exit 1
}
Write-Host "PASS"
exit 0
