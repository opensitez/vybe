# vybe-test: powershell/classes_sealed_members_and_classes/sealed_class_semantics_9
class FinalClass_9 {
    [int]$Val = 9
    [int]Compute() { return $this.Val * 2 }
}
$inst = [FinalClass_9]::new()
if ($inst.Compute() -ne (9 * 2)) { Write-Host "FAIL: Final class failed"; exit 1 }
Write-Host "PASS"; exit 0
