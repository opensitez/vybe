# vybe-test: powershell/convertfrom_markdown_cmdlet/markdown_heading_h1_conversion
# ConvertFrom-Markdown transforms '#' top-level markdown header into HTML <h1>
$info = ConvertFrom-Markdown -InputObject "# Document Header"

if ($null -eq $info) {
    Write-Host "FAIL: ConvertFrom-Markdown returned `$null"
    exit 1
}

if ($info.Html -notmatch "<h1.*>Document Header</h1>") {
    Write-Host "FAIL: expected <h1>Document Header</h1>, got: '$($info.Html)'"
    exit 1
}

Write-Host "PASS"
exit 0
