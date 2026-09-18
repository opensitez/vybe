# vybe-test: powershell/convertfrom_markdown_cmdlet/markdown_inline_bold_conversion
# ConvertFrom-Markdown transforms '**text**' into HTML <strong>text</strong>
$info = ConvertFrom-Markdown -InputObject "This is **strong text** in paragraph."

if ($info.Html -notmatch "<strong>strong text</strong>") {
    Write-Host "FAIL: expected <strong>strong text</strong>, got: '$($info.Html)'"
    exit 1
}

Write-Host "PASS"
exit 0
