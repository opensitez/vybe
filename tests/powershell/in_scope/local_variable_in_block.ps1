# vybe-test: powershell/in_scope/local_variable_in_block
{
    $x = 1
}
if ($x -ne $null) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
