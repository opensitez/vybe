# vybe-test: powershell/pipeline/select_object_property
$objects = @(
    [PSCustomObject]@{ Name = "Alice"; Age = 30 },
    [PSCustomObject]@{ Name = "Bob"; Age = 25 }
)
$names = $objects | Select-Object -ExpandProperty Name
if ($names.Count -ne 2) {
    Write-Host "FAIL: expected 2 names, got $($names.Count)"
    exit 1
}
Write-Host "PASS"
exit 0
