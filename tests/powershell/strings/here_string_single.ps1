# vybe-test: powershell/strings/here_string_single
$text = @'
Line 1
Line 2
'@
$lines = $text -split "`n"
if ($lines.Count -lt 2) {
    Write-Host "FAIL: expected at least 2 lines, got $($lines.Count)"
    exit 1
}
Write-Host "PASS"
exit 0
