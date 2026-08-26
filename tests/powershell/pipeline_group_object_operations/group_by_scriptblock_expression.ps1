# vybe-test: powershell/pipeline_group_object_operations/group_by_scriptblock_expression
$numbers = @(1, 2, 3, 4, 5, 6, 7, 8)
$groups = @($numbers | Group-Object { if ($_ % 2 -eq 0) { "Even" } else { "Odd" } })
if ($groups.Count -ne 2) {
    Write-Host "FAIL: Group-Object scriptblock failed"
    exit 1
}
$evenGroup = $groups | Where-Object { $_.Name -eq "Even" }
$oddGroup = $groups | Where-Object { $_.Name -eq "Odd" }
if ($evenGroup.Count -ne 4 -or $oddGroup.Count -ne 4) {
    Write-Host "FAIL: Even/Odd group counts failed"
    exit 1
}
Write-Host "PASS"
exit 0
