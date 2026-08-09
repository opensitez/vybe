# vybe-test: powershell/using_variable_scope/using_variable_hashtable_lookup
$dict = @{ Target = "Found" }
$sb = { ($using:dict)["Target"] }
$res = &$sb
if ($res -ne "Found") {
    Write-Host "FAIL: hashtable lookup (\$using:dict)['Target'] expected 'Found', got '$res'"
    exit 1
}
Write-Host "PASS"
exit 0
