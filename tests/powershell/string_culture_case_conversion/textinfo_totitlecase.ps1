# vybe-test: powershell/string_culture_case_conversion/textinfo_totitlecase
$ci = [System.Globalization.CultureInfo]::InvariantCulture
$title = $ci.TextInfo.ToTitleCase("war and peace")
if ($title -ne "War And Peace") {
    Write-Host "FAIL: ToTitleCase failed, got $title"
    exit 1
}
Write-Host "PASS"
exit 0
