# vybe-test: powershell/convertfrom_markdown_cmdlet/markdown_link_hyperlink_conversion
# ConvertFrom-Markdown transforms markdown links [Text](URL) into <a href="URL">Text</a>
$info = ConvertFrom-Markdown -InputObject '[Microsoft Docs](https://docs.microsoft.com)'

if ($info.Html -notmatch '<a href="https://docs\.microsoft\.com">Microsoft Docs</a>') {
    Write-Host "FAIL: anchor tag conversion mismatch, got: '$($info.Html)'"
    exit 1
}

Write-Host "PASS"
exit 0
