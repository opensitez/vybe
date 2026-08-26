# vybe-test: powershell/classes_sealed_members_and_classes/sealed_class_semantics_1
class FinalClass_1 {
    [int]$Val = 1
    [int]Compute() { return $this.Val * 2 }
}
$inst = [FinalClass_1]::new()
if ($inst.Compute() -ne (1 * 2)) { Write-Host "FAIL: Final class failed"; exit 1 }
Write-Host "PASS"; exit 0
