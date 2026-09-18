# vybe-test: powershell/select_string_cmdlet/select_string_simplematch_disables_regex
# -SimpleMatch treats special regex characters (like $, ., *) as literal characters
$line = "Subscription price is $19.99 per month."
$match = $line | Select-String -Pattern "$19.99" -SimpleMatch

if ($null -eq $match) {
    Write-Host "FAIL: Select-String -SimpleMatch failed to match literal '$19.99'"
    exit 1
}

if ($match.Line -ne $line) {
    Write-Host "FAIL: matched line mismatch, got: '$($match.Line)'"
    exit 1
}

Write-Host "PASS"
exit 0
