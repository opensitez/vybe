# vybe-test: powershell/using_variable_scope/using_variable_curly_braces
${spaced var} = "SpaceVal"
$sb = { ${using:spaced var} }
$res = &$sb
if ($res -ne "SpaceVal") {
    Write-Host "FAIL: \${using:spaced var} expected 'SpaceVal', got '$res'"
    exit 1
}
Write-Host "PASS"
exit 0
