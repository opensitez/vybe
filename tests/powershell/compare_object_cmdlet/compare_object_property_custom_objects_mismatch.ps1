# vybe-test: powershell/compare_object_cmdlet/compare_object_property_custom_objects_mismatch
# -Property surfaces differences when the specified property values differ
$obj1 = [pscustomobject]@{ UserId = 101; Active = $true }
$obj2 = [pscustomobject]@{ UserId = 202; Active = $true }

$diff = Compare-Object @($obj1) @($obj2) -Property UserId

if ($diff.Count -ne 2) {
    Write-Host "FAIL: expected 2 diff records, got $($diff.Count)"
    exit 1
}

$uids = @($diff | ForEach-Object { $_.UserId })
if ($uids -notcontains 101 -or $uids -notcontains 202) {
    Write-Host "FAIL: expected UserIds 101 and 202, got: @($($uids -join ', '))"
    exit 1
}

Write-Host "PASS"
exit 0
