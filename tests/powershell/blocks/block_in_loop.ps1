# vybe-test: powershell/blocks/block_in_loop
for ($i=0; $i -lt 1; $i++) {
    {
        Write-Output 'PASS'
    }
}
exit 0
