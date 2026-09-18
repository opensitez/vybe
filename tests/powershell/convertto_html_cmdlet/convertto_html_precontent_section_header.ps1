# vybe-test: powershell/convertto_html_cmdlet/convertto_html_precontent_section_header
# The -PreContent parameter prepends content before the main output table
$pre = "<h2>Production Cluster Status</h2>"
$html = [pscustomobject]@{ Cluster = "EU-West" } | ConvertTo-Html -PreContent $pre -Fragment
$text = $html -join " "

if ($text -notmatch [regex]::Escape($pre)) {
    Write-Host "FAIL: PreContent missing from HTML output"
    exit 1
}

Write-Host "PASS"
exit 0
