# vybe-test: powershell/convertfrom_markdown_cmdlet/markdown_heading_id_slug_generation
# ConvertFrom-Markdown generates URL slugs for heading id attributes
$info = ConvertFrom-Markdown -InputObject "# Advanced Configuration Guide"

if ($info.Html -notmatch 'id="advanced-configuration-guide"') {
    Write-Host "FAIL: heading ID slug mismatch, got: '$($info.Html)'"
    exit 1
}

Write-Host "PASS"
exit 0
