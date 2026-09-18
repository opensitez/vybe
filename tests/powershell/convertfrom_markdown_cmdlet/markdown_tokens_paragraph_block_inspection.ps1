# vybe-test: powershell/convertfrom_markdown_cmdlet/markdown_tokens_paragraph_block_inspection
# The .Tokens collection contains ParagraphBlock tokens for regular text paragraphs
$info = ConvertFrom-Markdown -InputObject "This is a simple paragraph."

$pTok = $info.Tokens | Where-Object { $_.GetType().Name -eq "ParagraphBlock" }

if ($null -eq $pTok) {
    Write-Host "FAIL: ParagraphBlock token missing in .Tokens collection"
    exit 1
}

Write-Host "PASS"
exit 0
