# vybe-test: powershell/in_scope/private_scope
function Test-Func {
    $private:x = 1
}
Test-Func
Write-Host 'PASS'
exit 0
