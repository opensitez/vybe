# vybe-test: powershell/variables/variable_dynamic_name
$varName = "dynamic"
Set-Variable -Name $varName -Value 42
$result = Get-Variable -Name $varName -ValueOnly
if ($result -ne 42) {
    Write-Host "FAIL: expected 42, got $result"
    exit 1
}
# Also via variable: drive
$val = (Get-Item "variable:$varName").Value
if ($val -ne 42) { Write-Host "FAIL: variable: drive"; exit 1 }
Write-Host "PASS"
exit 0
