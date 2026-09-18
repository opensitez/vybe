# vybe-test: powershell/convertfrom_markdown_cmdlet/markdown_tokens_heading_block_inspection
# The .Tokens collection contains Markdig syntax AST blocks including HeadingBlock
$info = ConvertFrom-Markdown -InputObject "# Heading Only"

$headingTok = $info.Tokens | Where-Object { $_.GetType().Name -eq "HeadingBlock" }

if ($null -eq $headingTok) {
    Write-Host "FAIL: HeadingBlock token missing in .Tokens collection"
    exit 1
}

if ($headingTok.Level -ne 1) {
    Write-Host "FAIL: expected HeadingBlock.Level 1, got: $($headingTok.Level)"
    exit 1
}

Write-Host "PASS"
exit 0
