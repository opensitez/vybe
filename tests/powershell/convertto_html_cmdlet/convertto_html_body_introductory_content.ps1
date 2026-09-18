# vybe-test: powershell/convertto_html_cmdlet/convertto_html_body_introductory_content
# The -Body parameter inserts content immediately preceding the table in <body>
$bodyContent = "<p>Introductory overview paragraph.</p>"
$html = [pscustomobject]@{ Item = "Database" } | ConvertTo-Html -Body $bodyContent
$text = $html -join " "

if ($text -notmatch [regex]::Escape($bodyContent)) {
    Write-Host "FAIL: body content missing from output"
    exit 1
}

Write-Host "PASS"
exit 0
