# vybe-test: powershell/comment_syntax_suite/comment_only_statement_block
{
    # no-op statement block
}

$ran = $true
if (-not $ran) {
    Write-Host 'FAIL: execution never reached after empty comment block'
    exit 1
}

Write-Host 'PASS'
exit 0
