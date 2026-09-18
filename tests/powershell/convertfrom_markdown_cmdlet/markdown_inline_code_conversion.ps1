# vybe-test: powershell/convertfrom_markdown_cmdlet/markdown_inline_code_conversion
# ConvertFrom-Markdown transforms backtick-enclosed words into HTML <code>
$bt = [char]96
$markdown = "Run " + $bt + "Get-Process" + $bt + " now."

$info = ConvertFrom-Markdown -InputObject $markdown

if ($info.Html -notmatch "<code>Get-Process</code>") {
    Write-Host "FAIL: inline code tag missing in HTML: '$($info.Html)'"
    exit 1
}

Write-Host "PASS"
exit 0
