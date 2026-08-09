# vybe-test: powershell/object_sorting/sort_single_property
$items = @([pscustomobject]@{ Age = 30 }, [pscustomobject]@{ Age = 20 }, [pscustomobject]@{ Age = 25 })
$res = $items | Sort-Object -Property Age
if ($res[0].Age -ne 20 -or $res[2].Age -ne 30) {
    Write-Host "FAIL: Sort-Object single property expected 20, 25, 30"
    exit 1
}
Write-Host "PASS"
exit 0
