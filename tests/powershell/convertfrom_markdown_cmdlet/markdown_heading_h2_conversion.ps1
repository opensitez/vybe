# vybe-test: powershell/convertfrom_markdown_cmdlet/markdown_heading_h2_conversion
# ConvertFrom-Markdown transforms '##' header into HTML <h2>
$info = ConvertFrom-Markdown -InputObject "## Subtitle Section"

if ($info.Html -notmatch "<h2.*>Subtitle Section</h2>") {
    Write-Host "FAIL: expected <h2>Subtitle Section</h2>, got: '$($info.Html)'"
    exit 1
}

Write-Host "PASS"
exit 0
