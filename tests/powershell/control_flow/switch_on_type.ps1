# vybe-test: powershell/control_flow/switch_on_type
$values = @(42, "hello", 3.14, $true)
$results = @()
foreach ($v in $values) {
    switch ($v) {
        { $_ -is [int] }    { $results += "int";    break }
        { $_ -is [string] } { $results += "string"; break }
        { $_ -is [double] } { $results += "double"; break }
        { $_ -is [bool] }   { $results += "bool";   break }
        default              { $results += "other" }
    }
}
if ($results[0] -ne "int")    { Write-Host "FAIL: int";    exit 1 }
if ($results[1] -ne "string") { Write-Host "FAIL: string"; exit 1 }
if ($results[2] -ne "double") { Write-Host "FAIL: double"; exit 1 }
if ($results[3] -ne "bool")   { Write-Host "FAIL: bool";   exit 1 }
Write-Host "PASS"
exit 0
