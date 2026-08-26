# vybe-test: powershell/classes_sealed_members_and_classes/sealed_class_semantics_17
class FinalClass_17 {
    [int]$Val = 17
    [int]Compute() { return $this.Val * 2 }
}
$inst = [FinalClass_17]::new()
if ($inst.Compute() -ne (17 * 2)) { Write-Host "FAIL: Final class failed"; exit 1 }
Write-Host "PASS"; exit 0
