# vybe-test: powershell/objects/pscustomobject_array_select
$people = @(
    [PSCustomObject]@{ Name = "Alice"; Age = 30 },
    [PSCustomObject]@{ Name = "Bob";   Age = 25 },
    [PSCustomObject]@{ Name = "Carol"; Age = 35 }
)
$names = $people | Select-Object -ExpandProperty Name
if ($names.Count -ne 3)    { Write-Host "FAIL: count"; exit 1 }
if ($names[0] -ne "Alice") { Write-Host "FAIL: [0]";   exit 1 }
$youngest = $people | Sort-Object Age | Select-Object -First 1
if ($youngest.Name -ne "Bob") { Write-Host "FAIL: youngest is '$($youngest.Name)'"; exit 1 }
Write-Host "PASS"
exit 0
