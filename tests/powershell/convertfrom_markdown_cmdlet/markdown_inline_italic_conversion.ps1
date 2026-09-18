# vybe-test: powershell/convertfrom_markdown_cmdlet/markdown_inline_italic_conversion
# ConvertFrom-Markdown transforms '*text*' into HTML <em>text</em>
$info = ConvertFrom-Markdown -InputObject "This is *emphasized text* in sentence."

if ($info.Html -notmatch "<em>emphasized text</em>") {
    Write-Host "FAIL: expected <em>emphasized text</em>, got: '$($info.Html)'"
    exit 1
}

Write-Host "PASS"
exit 0
