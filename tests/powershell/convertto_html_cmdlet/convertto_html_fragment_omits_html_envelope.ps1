# vybe-test: powershell/convertto_html_cmdlet/convertto_html_fragment_omits_html_envelope
# The -Fragment switch suppresses outer document tags and outputs only the HTML table fragment
$html = [pscustomobject]@{ Setting = "Optimized" } | ConvertTo-Html -Fragment
$text = $html -join " "

if ($text -match "<!DOCTYPE" -or $text -match "<html" -or $text -match "<body") {
    Write-Host "FAIL: document envelope tags found in -Fragment output"
    exit 1
}

if ($text -notmatch "<table>.*</table>") {
    Write-Host "FAIL: <table> tags missing from -Fragment output"
    exit 1
}

Write-Host "PASS"
exit 0
