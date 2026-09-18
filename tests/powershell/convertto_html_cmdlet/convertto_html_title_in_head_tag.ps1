# vybe-test: powershell/convertto_html_cmdlet/convertto_html_title_in_head_tag
# The -Title parameter populates the <title> element inside the HTML document <head>
$html = [pscustomobject]@{ MemoryUsage = 45 } | ConvertTo-Html -Title "Server Metrics Report"
$text = $html -join " "

if ($text -notmatch "<title>Server Metrics Report</title>") {
    Write-Host "FAIL: <title> tag does not contain expected title: $text"
    exit 1
}

Write-Host "PASS"
exit 0
