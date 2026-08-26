# vybe-test: powershell/classes_sealed_members_and_classes/sealed_class_semantics_15
class FinalClass_15 {
    [int]$Val = 15
    [int]Compute() { return $this.Val * 2 }
}
$inst = [FinalClass_15]::new()
if ($inst.Compute() -ne (15 * 2)) { Write-Host "FAIL: Final class failed"; exit 1 }
Write-Host "PASS"; exit 0
