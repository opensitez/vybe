# vybe-test: powershell/strings/escape_sequences
$tab = "Hello`tWorld"
$result = $tab.Contains("`t")
if ($result -ne $true) {
    Write-Host "FAIL: expected tab character in string"
    exit 1
}
Write-Host "PASS"
exit 0
