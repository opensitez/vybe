# vybe-test: powershell/pipeline_group_object_operations/group_object_with_datetime_year_scriptblock
$d1 = [datetime]::Parse("2024-01-01")
$d2 = [datetime]::Parse("2024-06-01")
$d3 = [datetime]::Parse("2025-01-01")
$groups = @($d1, $d2, $d3 | Group-Object { $_.Year })
if ($groups.Count -ne 2) {
    Write-Host "FAIL: Group-Object datetime year scriptblock failed"
    exit 1
}
Write-Host "PASS"
exit 0
