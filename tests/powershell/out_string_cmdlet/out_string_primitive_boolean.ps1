# vybe-test: powershell/out_string_cmdlet/out_string_primitive_boolean
# Booleans convert to their string representations 'True' and 'False'
$outTrue  = $true | Out-String -NoNewline
$outFalse = $false | Out-String -NoNewline

if ($outTrue -ne "True" -or $outFalse -ne "False") {
    Write-Host "FAIL: boolean string conversions mismatch: true='$outTrue', false='$outFalse'"
    exit 1
}

Write-Host "PASS"
exit 0
