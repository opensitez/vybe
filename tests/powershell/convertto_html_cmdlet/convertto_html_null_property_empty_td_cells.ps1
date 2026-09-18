# vybe-test: powershell/convertto_html_cmdlet/convertto_html_null_property_empty_td_cells
# ConvertTo-Html represents $null property values as empty <td></td> table cells
$obj = [pscustomobject]@{ Exists = "Present"; Missing = $null }
$html = $obj | ConvertTo-Html -Fragment
$text = $html -join " "

if ($text -notmatch "<td>Present</td><td></td>") {
    Write-Host "FAIL: null property was not rendered as empty <td></td> cell: $text"
    exit 1
}

Write-Host "PASS"
exit 0
