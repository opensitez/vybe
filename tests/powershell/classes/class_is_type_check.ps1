# vybe-test: powershell/classes/class_is_type_check
class Fruit {}
class Apple : Fruit {}
$a = [Apple]::new()
if ($a -isnot [Apple]) { Write-Host "FAIL: not Apple"; exit 1 }
if ($a -isnot [Fruit]) { Write-Host "FAIL: not Fruit (base)"; exit 1 }
if ($a -is [System.Object]) {} else { Write-Host "FAIL: not Object"; exit 1 }
Write-Host "PASS"
exit 0
