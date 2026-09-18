# vybe-test: powershell/convertfrom_markdown_cmdlet/markdown_unordered_list_conversion
# ConvertFrom-Markdown transforms bullet points into <ul> and <li> tags
$listText = "- Alpha item" + [char]10 + "- Beta item"
$info = ConvertFrom-Markdown -InputObject $listText

if ($info.Html -notmatch "<ul>") {
    Write-Host "FAIL: <ul> tag missing in unordered list HTML: '$($info.Html)'"
    exit 1
}

if ($info.Html -notmatch "<li>Alpha item</li>" -or $info.Html -notmatch "<li>Beta item</li>") {
    Write-Host "FAIL: <li> elements missing in HTML: '$($info.Html)'"
    exit 1
}

Write-Host "PASS"
exit 0
