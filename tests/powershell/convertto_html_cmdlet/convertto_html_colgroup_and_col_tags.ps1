# vybe-test: powershell/convertto_html_cmdlet/convertto_html_colgroup_and_col_tags
# ConvertTo-Html emits valid XHTML colgroup and col elements matching column count
$obj = [pscustomobject]@{ Col1 = "A"; Col2 = "B" }
$html = $obj | ConvertTo-Html -Fragment
$text = $html -join " "

if ($text -notmatch "<colgroup><col/><col/></colgroup>") {
    Write-Host "FAIL: colgroup element mismatch for 2-column table: $text"
    exit 1
}

Write-Host "PASS"
exit 0
