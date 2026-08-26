# vybe-test: powershell/csv_header_manipulation/header_renaming_via_calculated_properties
$csv = @"
OldName,OldVal
Item1,100
"@
$renamed = $csv | ConvertFrom-Csv | Select-Object @{ N = "NewName"; E = { $_.OldName } }, @{ N = "NewVal"; E = { $_.OldVal } }
if ($renamed.NewName -ne "Item1" -or $renamed.NewVal -ne "100") {
    Write-Host "FAIL: CSV header renaming via calculated properties failed"
    exit 1
}
Write-Host "PASS"
exit 0
