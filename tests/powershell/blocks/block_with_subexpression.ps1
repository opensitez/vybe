# vybe-test: powershell/blocks/block_with_subexpression
{
    $(1 + 1) | Out-Null
    Write-Output 'PASS'
}
exit 0
