# vybe-test: powershell/blocks/block_with_expression
{
    1 + 1 | Out-Null
    Write-Output 'PASS'
}
exit 0
