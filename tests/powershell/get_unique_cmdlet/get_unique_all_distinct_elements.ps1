# vybe-test: powershell/get_unique_cmdlet/get_unique_all_distinct_elements
# A collection with all distinct elements retains all items without change
$items = @(10, 20, 30, 40)
$res = @($items | Get-Unique)

if ($res.Count -ne 4) {
    Write-Host "FAIL: all distinct items should be preserved, got $($res.Count)"
    exit 1
}

Write-Host "PASS"
exit 0
