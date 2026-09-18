# vybe-test: powershell/convertto_html_cmdlet/convertto_html_boolean_and_numeric_rendering
# ConvertTo-Html correctly formats boolean and integer primitive properties
$obj = [pscustomobject]@{ IsAdmin = $true; Attempts = 0 }
$html = $obj | ConvertTo-Html -Fragment
$text = $html -join " "

if ($text -notmatch "<td>True</td><td>0</td>") {
    Write-Host "FAIL: boolean and numeric formatting mismatch, got: $text"
    exit 1
}

Write-Host "PASS"
exit 0
