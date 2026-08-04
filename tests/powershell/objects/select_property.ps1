# vybe-test: powershell/objects/select_property
$obj = [PSCustomObject]@{ Name = "Alice"; Age = 25; City = "NYC" }
$selected = $obj | Select-Object Name, Age
$hasCity = $null -ne ($selected.PSObject.Properties | Where-Object { $_.Name -eq "City" })
if ($hasCity -eq $true) {
    Write-Host "FAIL: City property should not be selected"
    exit 1
}
if ($selected.Name -ne "Alice") {
    Write-Host "FAIL: expected Name to be Alice"
    exit 1
}
Write-Host "PASS"
exit 0
