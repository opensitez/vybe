# vybe-test: powershell/convertfrom_markdown_cmdlet/markdown_horizontal_rule_conversion
# ConvertFrom-Markdown transforms '---' into HTML <hr />
$info = ConvertFrom-Markdown -InputObject "---"

if ($info.Html -notmatch "<hr />") {
    Write-Host "FAIL: expected <hr /> tag, got: '$($info.Html)'"
    exit 1
}

Write-Host "PASS"
exit 0
