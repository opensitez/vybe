# vybe-test: powershell/convertto_html_cmdlet/convertto_html_inputobject_parameter_binding
# ConvertTo-Html binds arguments supplied directly to -InputObject
$data = [pscustomobject]@{ ResponseCode = 200 }
$html = ConvertTo-Html -InputObject $data -Fragment
$text = $html -join " "

if ($text -notmatch "<th>ResponseCode</th>" -or $text -notmatch "<td>200</td>") {
    Write-Host "FAIL: -InputObject binding did not produce expected HTML row: $text"
    exit 1
}

Write-Host "PASS"
exit 0
