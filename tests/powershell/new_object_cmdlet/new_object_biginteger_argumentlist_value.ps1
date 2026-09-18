# vybe-test: powershell/new_object_cmdlet/new_object_biginteger_argumentlist_value
# New-Object System.Numerics.BigInteger -ArgumentList creates an arbitrary-precision integer
$bi = New-Object System.Numerics.BigInteger -ArgumentList 9999999999999

if ($bi.ToString() -ne "9999999999999") {
    Write-Host "FAIL: expected BigInteger '9999999999999', got: '$($bi.ToString())'"
    exit 1
}

Write-Host "PASS"
exit 0
