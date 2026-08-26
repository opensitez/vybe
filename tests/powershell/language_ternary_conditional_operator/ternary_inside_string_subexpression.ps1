# vybe-test: powershell/language_ternary_conditional_operator/ternary_inside_string_subexpression
$isAdmin = $true
$msg = "User role is: $( $isAdmin ? 'Administrator' : 'StandardUser' )"
if ($msg -ne "User role is: Administrator") {
    Write-Host "FAIL: Ternary in string subexpression failed, got '$msg'"
    exit 1
}
Write-Host "PASS"
exit 0
