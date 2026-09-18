# vybe-test: powershell/convertfrom_markdown_cmdlet/markdown_strikethrough_extension
# ConvertFrom-Markdown supports ~~text~~ strikethrough converting into <del>text</del>
$info = ConvertFrom-Markdown -InputObject "This feature is ~~deprecated~~ and removed."

if ($info.Html -notmatch "<del>deprecated</del>") {
    Write-Host "FAIL: strikethrough <del> tag missing in HTML: '$($info.Html)'"
    exit 1
}

Write-Host "PASS"
exit 0
