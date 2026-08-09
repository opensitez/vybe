# vybe-test: powershell/pipeline_chaining/chain_parentheses_grouping
$res = (($false || $true) && "GroupedPass")
if ($res -ne "GroupedPass") {
    Write-Host "FAIL: grouped chain expected 'GroupedPass', got '$res'"
    exit 1
}
Write-Host "PASS"
exit 0
