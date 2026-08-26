# vybe-test: powershell/classes_property_attributes/validateset_case_insensitive_behavior
class FruitBox {
    [ValidateSet("Apple", "Banana")][string]$Fruit
}
$fb = [FruitBox]::new()
$fb.Fruit = "apple" # case-insensitive valid match
if ($fb.Fruit -ne "apple" -and $fb.Fruit -ne "Apple") {
    Write-Host "FAIL: ValidateSet case-insensitive behavior failed"
    exit 1
}
Write-Host "PASS"
exit 0
