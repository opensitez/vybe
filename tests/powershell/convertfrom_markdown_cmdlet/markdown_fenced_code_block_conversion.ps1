# vybe-test: powershell/convertfrom_markdown_cmdlet/markdown_fenced_code_block_conversion
# ConvertFrom-Markdown transforms fenced code blocks into <pre><code class="language-...">
$fenced = '```powershell' + [char]10 + '$count = 42' + [char]10 + '```'
$info = ConvertFrom-Markdown -InputObject $fenced

if ($info.Html -notmatch '<pre><code class="language-powershell">') {
    Write-Host "FAIL: fenced code block HTML missing expected pre/code language tags: '$($info.Html)'"
    exit 1
}

Write-Host "PASS"
exit 0
