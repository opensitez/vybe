# vybe-test: powershell/type_nullable_value_types/type_checking_is_nullable
$t = [type]"System.Nullable[int]"
if ($t.IsGenericType -ne $true -or $t.GetGenericTypeDefinition() -ne [type]"System.Nullable``1") {
    Write-Host "FAIL: -is [Nullable[int]] check failed"
    exit 1
}
Write-Host "PASS"
exit 0
