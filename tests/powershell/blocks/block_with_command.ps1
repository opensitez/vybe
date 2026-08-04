# vybe-test: powershell/blocks/block_with_command
{
    Get-Command Write-Output | Out-Null
    Write-Output 'PASS'
}
exit 0
