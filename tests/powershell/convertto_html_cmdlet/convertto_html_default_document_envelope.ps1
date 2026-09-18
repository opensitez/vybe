# vybe-test: powershell/convertto_html_cmdlet/convertto_html_default_document_envelope
# ConvertTo-Html generates a complete XHTML document envelope by default
$html = [pscustomobject]@{ Status = "Active" } | ConvertTo-Html
$text = $html -join " "

if ($text -notmatch "<!DOCTYPE" -or $text -notmatch "<html" -or $text -notmatch "<head>" -or $text -notmatch "<body") {
    Write-Host "FAIL: document envelope elements missing from default ConvertTo-Html output"
    exit 1
}

Write-Host "PASS"
exit 0
