# vybe-test: powershell/null_coalescing_assignment/null_assignment_in_loop
$items = @($null, "Val1", $null)
$results = @()
foreach ($item in $items) {
    $x = $item
    $x ??= "FallbackVal"
    $results += $x
}
if ($results[0] -ne "FallbackVal" -or $results[1] -ne "Val1" -or $results[2] -ne "FallbackVal") {
    Write-Host "FAIL: loop ??= expected FallbackVal, Val1, FallbackVal"
    exit 1
}
Write-Host "PASS"
exit 0
