# vybe-test: powershell/generic_types/generic_type_definition_check
$t = [System.Collections.Generic.List[string]]
if (-not $t.IsGenericType) {
    Write-Host "FAIL: IsGenericType expected true"
    exit 1
}
$args = $t.GetGenericArguments()
if ($args[0] -ne [string]) {
    Write-Host "FAIL: Generic argument expected [string], got $($args[0])"
    exit 1
}
Write-Host "PASS"
exit 0
