# vybe-test: powershell/constant_variables/constant_variable_subexpression
New-Variable -Name "SUB_CONST" -Value 33 -Option Constant
$str = "Val: $( $SUB_CONST )"
if ($str -ne "Val: 33") {
    Write-Host "FAIL: Constant variable subexpression expected 'Val: 33', got '$str'"
    exit 1
}
Write-Host "PASS"
exit 0
