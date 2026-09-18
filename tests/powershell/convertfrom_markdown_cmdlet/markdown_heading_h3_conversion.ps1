# vybe-test: powershell/convertfrom_markdown_cmdlet/markdown_heading_h3_conversion
# ConvertFrom-Markdown transforms '###' header into HTML <h3>
$info = ConvertFrom-Markdown -InputObject "### Sub-Section Item"

if ($info.Html -notmatch "<h3.*>Sub-Section Item</h3>") {
    Write-Host "FAIL: expected <h3>Sub-Section Item</h3>, got: '$($info.Html)'"
    exit 1
}

Write-Host "PASS"
exit 0
