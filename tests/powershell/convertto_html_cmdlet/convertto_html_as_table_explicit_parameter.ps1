# vybe-test: powershell/convertto_html_cmdlet/convertto_html_as_table_explicit_parameter
# Passing -As Table explicitly produces tabular layout matching standard ConvertTo-Html behavior
$obj = [pscustomobject]@{ SettingName = "Threshold"; Value = 95 }
$html = $obj | ConvertTo-Html -As Table -Fragment
$text = $html -join " "

if ($text -notmatch "<th>SettingName</th><th>Value</th>") {
    Write-Host "FAIL: table headers missing in -As Table output"
    exit 1
}

if ($text -notmatch "<td>Threshold</td><td>95</td>") {
    Write-Host "FAIL: table data cells missing in -As Table output"
    exit 1
}

Write-Host "PASS"
exit 0
