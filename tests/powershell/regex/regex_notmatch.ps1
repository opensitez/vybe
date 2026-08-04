# vybe-test: powershell/regex/regex_notmatch
$str = "Hello World"
if ($str -notmatch "\d+") {
    Write-Host "PASS"
    exit 0
} else {
    Write-Host "FAIL: string should not match digit pattern"
    exit 1
}
