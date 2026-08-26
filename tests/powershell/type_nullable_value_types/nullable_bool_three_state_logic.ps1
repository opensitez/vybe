# vybe-test: powershell/type_nullable_value_types/nullable_bool_three_state_logic
$t = [type]"System.Nullable[bool]"
$instT = [System.Activator]::CreateInstance($t, @($true))
$instF = [System.Activator]::CreateInstance($t, @($false))
$valT = $t.GetProperty("Value").GetValue($instT)
$valF = $t.GetProperty("Value").GetValue($instF)
if ($valT -ne $true -or $valF -ne $false) {
    Write-Host "FAIL: Three-state boolean nullable failed"
    exit 1
}
Write-Host "PASS"
exit 0
