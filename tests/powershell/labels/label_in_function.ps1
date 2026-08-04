# vybe-test: powershell/labels/label_in_function
function Test-Func {
    :start
    Write-Output 'PASS'
}
Test-Func
exit 0
