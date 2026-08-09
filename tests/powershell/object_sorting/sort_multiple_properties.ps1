# vybe-test: powershell/object_sorting/sort_multiple_properties
$items = @(
    [pscustomobject]@{ Dept = "IT"; Name = "Bob" },
    [pscustomobject]@{ Dept = "HR"; Name = "Alice" },
    [pscustomobject]@{ Dept = "IT"; Name = "Adam" }
)
$res = $items | Sort-Object -Property Dept, Name
if ($res[0].Name -ne "Alice" -or $res[1].Name -ne "Adam" -or $res[2].Name -ne "Bob") {
    Write-Host "FAIL: Sort-Object multiple properties expected Alice (HR), Adam (IT), Bob (IT)"
    exit 1
}
Write-Host "PASS"
exit 0
