# vybe-test: powershell/token_and_grammar_system/comment_block_single
$a = 1
<# single block comment should be ignored #>
$b = 2
if (($a + $b) -ne 3) {
    Write-Host "FAIL: block comment affected parsing/evaluation incorrectly"
    exit 1
}

Write-Host 'PASS'
exit 0
