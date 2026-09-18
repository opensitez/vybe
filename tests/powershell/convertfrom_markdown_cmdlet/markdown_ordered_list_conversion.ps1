# vybe-test: powershell/convertfrom_markdown_cmdlet/markdown_ordered_list_conversion
# ConvertFrom-Markdown transforms numbered list into <ol> and <li> tags
$listText = "1. First step" + [char]10 + "2. Second step"
$info = ConvertFrom-Markdown -InputObject $listText

if ($info.Html -notmatch "<ol>") {
    Write-Host "FAIL: <ol> tag missing in ordered list HTML: '$($info.Html)'"
    exit 1
}

if ($info.Html -notmatch "<li>First step</li>" -or $info.Html -notmatch "<li>Second step</li>") {
    Write-Host "FAIL: <li> elements missing in ordered list: '$($info.Html)'"
    exit 1
}

Write-Host "PASS"
exit 0
