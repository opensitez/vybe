# vybe-test: powershell/pipeline_group_object_operations/group_object_with_empty_input
$groups = @(@() | Group-Object -Property Name)
if ($groups.Length -ne 0) {
    Write-Host "FAIL: Group-Object on empty input should return empty"
    exit 1
}
Write-Host "PASS"
exit 0
