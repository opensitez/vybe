# vybe-test: powershell/convertto_html_cmdlet/convertto_html_property_selection_columns
# The -Property parameter restricts generated table columns to specified properties
$obj = [pscustomobject]@{ VisibleProp = "Shown"; HiddenProp = "Secret" }
$html = $obj | ConvertTo-Html -Property VisibleProp -Fragment
$text = $html -join " "

if ($text -notmatch "<th>VisibleProp</th>") {
    Write-Host "FAIL: selected property column missing from table header"
    exit 1
}

if ($text -match "<th>HiddenProp</th>" -or $text -match "<td>Secret</td>") {
    Write-Host "FAIL: unselected property was included in table"
    exit 1
}

Write-Host "PASS"
exit 0
