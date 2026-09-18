# vybe-test: powershell/convertfrom_markdown_cmdlet/markdown_blockquote_conversion
# ConvertFrom-Markdown transforms lines starting with '>' into <blockquote>
$info = ConvertFrom-Markdown -InputObject "> Important notice: system reboot tonight."

if ($info.Html -notmatch "<blockquote>") {
    Write-Host "FAIL: blockquote tag missing in HTML: '$($info.Html)'"
    exit 1
}

if ($info.Html -notmatch "Important notice: system reboot tonight\.") {
    Write-Host "FAIL: blockquote inner text missing in HTML: '$($info.Html)'"
    exit 1
}

Write-Host "PASS"
exit 0
