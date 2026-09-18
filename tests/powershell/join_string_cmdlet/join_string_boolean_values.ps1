# vybe-test: powershell/join_string_cmdlet/join_string_boolean_values
$bools = @($true, $false, $true)

# Boolean values format as True and False
$res = $bools | Join-String -Separator '/'

if ($res -ne "True/False/True") {
    Write-Host "FAIL: expected 'True/False/True', got: '$res'"
    exit 1
}

Write-Host "PASS"
exit 0
