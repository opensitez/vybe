# vybe-test: powershell/convertto_html_cmdlet/convertto_html_as_list_property_table
# -As List renders properties as alternating key-value cell pairs inside a vertical table
$html = [pscustomobject]@{ Hostname = "node-01"; Port = 8080 } | ConvertTo-Html -As List -Fragment
$text = $html -join " "

if ($text -notmatch "<td>Hostname:</td><td>node-01</td>") {
    Write-Host "FAIL: Hostname key-value row missing in list view"
    exit 1
}

if ($text -notmatch "<td>Port:</td><td>8080</td>") {
    Write-Host "FAIL: Port key-value row missing in list view"
    exit 1
}

Write-Host "PASS"
exit 0
