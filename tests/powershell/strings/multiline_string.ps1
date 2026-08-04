# vybe-test: powershell/strings/multiline_string
$text = "Line 1
Line 2
Line 3"
$lines = $text -split "`n"
if ($lines.Count -lt 3) {
    Write-Host "FAIL: expected at least 3 lines"
    exit 1
}
Write-Host "PASS"
exit 0
