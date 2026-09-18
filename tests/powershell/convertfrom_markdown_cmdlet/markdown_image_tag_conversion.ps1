# vybe-test: powershell/convertfrom_markdown_cmdlet/markdown_image_tag_conversion
# ConvertFrom-Markdown transforms image markdown ![Alt](Src) into <img src="Src" alt="Alt" />
$info = ConvertFrom-Markdown -InputObject '![Company Logo](assets/logo.png)'

if ($info.Html -notmatch '<img src="assets/logo\.png" alt="Company Logo"') {
    Write-Host "FAIL: image tag conversion mismatch, got: '$($info.Html)'"
    exit 1
}

Write-Host "PASS"
exit 0
