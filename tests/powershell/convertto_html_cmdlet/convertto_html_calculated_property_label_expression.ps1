# vybe-test: powershell/convertto_html_cmdlet/convertto_html_calculated_property_label_expression
# ConvertTo-Html evaluates calculated properties specified via @{Label=...; Expression=...} hashtables
$calc = @{
    Label = "Squared"
    Expression = { $_.Number * $_.Number }
}

$html = [pscustomobject]@{ Number = 7 } | ConvertTo-Html -Property $calc -Fragment
$text = $html -join " "

if ($text -notmatch "<th>Squared</th>") {
    Write-Host "FAIL: calculated property header 'Squared' missing"
    exit 1
}

if ($text -notmatch "<td>49</td>") {
    Write-Host "FAIL: calculated property expression result 49 missing"
    exit 1
}

Write-Host "PASS"
exit 0
