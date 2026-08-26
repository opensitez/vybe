# vybe-test: powershell/pipeline_group_object_operations/group_object_notitle_noelement_flag
$items = @("apple", "apricot", "banana")
$groups = @($items | Group-Object { $_.Substring(0,1) } -NoElement)
if ($groups.Count -ne 2 -or $groups[0].Group -ne $null) {
    Write-Host "FAIL: Group-Object -NoElement should omit Group elements"
    exit 1
}
Write-Host "PASS"
exit 0
