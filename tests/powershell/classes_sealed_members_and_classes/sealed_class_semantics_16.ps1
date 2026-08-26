# vybe-test: powershell/classes_sealed_members_and_classes/sealed_class_semantics_16
class FinalClass_16 {
    [int]$Val = 16
    [int]Compute() { return $this.Val * 2 }
}
$inst = [FinalClass_16]::new()
if ($inst.Compute() -ne (16 * 2)) { Write-Host "FAIL: Final class failed"; exit 1 }
Write-Host "PASS"; exit 0
