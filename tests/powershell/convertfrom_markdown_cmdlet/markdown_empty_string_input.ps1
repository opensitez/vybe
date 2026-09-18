# vybe-test: powershell/convertfrom_markdown_cmdlet/markdown_empty_string_input
# An empty string input yields an empty HTML string without throwing
$info = ConvertFrom-Markdown -InputObject ""

if ($null -eq $info) {
    Write-Host "FAIL: ConvertFrom-Markdown returned `$null for empty string"
    exit 1
}

if ($info.Html -ne "") {
    Write-Host "FAIL: expected empty HTML output for empty input, got: '$($info.Html)'"
    exit 1
}

Write-Host "PASS"
exit 0
