# vybe-test: powershell/blocks/block_in_function
function Test-Func {
    {
        Write-Output 'PASS'
    }
}
Test-Func
exit 0
