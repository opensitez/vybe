# vybe-test: powershell/convertto_html_cmdlet/convertto_html_nested_custom_object_rendering
# ConvertTo-Html stringifies nested PSCustomObjects into @{Key=Val} representations
$obj = [pscustomobject]@{
    Config = [pscustomobject]@{ Port = 8080 }
}

$html = $obj | ConvertTo-Html -Fragment
$text = $html -join " "

if ($text -notmatch "<td>@\{Port=8080\}</td>") {
    Write-Host "FAIL: nested object stringification mismatch: $text"
    exit 1
}

Write-Host "PASS"
exit 0
