# vybe-test: powershell/parameters_validate_set/validateset_with_char_type_parameter
function Select-Char {
    param([ValidateSet("Y", "N")][char]$Choice)
    return $Choice
}
$res = Select-Char -Choice ([char]'Y')
if ($res -ne [char]'Y') {
    Write-Host "FAIL: Char parameter with ValidateSet failed"
    exit 1
}
Write-Host "PASS"
exit 0
