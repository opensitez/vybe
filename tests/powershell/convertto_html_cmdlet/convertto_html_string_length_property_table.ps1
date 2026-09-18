# vybe-test: powershell/convertto_html_cmdlet/convertto_html_string_length_property_table
# Piping string primitives into ConvertTo-Html reflects the Length property in table columns
$html = "PowerShell" | ConvertTo-Html -Fragment
$text = $html -join " "

if ($text -notmatch "<th>Length</th>") {
    Write-Host "FAIL: Length header missing for string primitive input"
    exit 1
}

if ($text -notmatch "<td>10</td>") {
    Write-Host "FAIL: string length value 10 missing from table row"
    exit 1
}

Write-Host "PASS"
exit 0
