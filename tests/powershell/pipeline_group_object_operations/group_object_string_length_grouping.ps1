# vybe-test: powershell/pipeline_group_object_operations/group_object_string_length_grouping
$words = @("hi", "to", "cat", "dog", "elephant")
$groups = @($words | Group-Object -Property Length)
if ($groups.Count -ne 3) { # length 2 (hi, to), length 3 (cat, dog), length 8 (elephant)
    Write-Host "FAIL: Group-Object string length grouping failed"
    exit 1
}
Write-Host "PASS"
exit 0
