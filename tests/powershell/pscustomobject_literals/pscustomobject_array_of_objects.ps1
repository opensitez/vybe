# vybe-test: powershell/pscustomobject_literals/pscustomobject_array_of_objects
$list = @(
    [pscustomobject]@{ Val = 10 }
    [pscustomobject]@{ Val = 20 }
)
$sum = ($list | Measure-Object -Property Val -Sum).Sum
if ($sum -ne 30) {
    Write-Host "FAIL: sum expected 30, got $sum"
    exit 1
}
Write-Host "PASS"
exit 0
