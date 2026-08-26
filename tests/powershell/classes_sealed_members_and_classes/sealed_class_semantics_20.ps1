# vybe-test: powershell/classes_sealed_members_and_classes/sealed_class_semantics_20
class FinalClass_20 {
    [int]$Val = 20
    [int]Compute() { return $this.Val * 2 }
}
$inst = [FinalClass_20]::new()
if ($inst.Compute() -ne (20 * 2)) { Write-Host "FAIL: Final class failed"; exit 1 }
Write-Host "PASS"; exit 0
