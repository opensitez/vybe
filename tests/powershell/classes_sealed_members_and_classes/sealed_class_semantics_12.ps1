# vybe-test: powershell/classes_sealed_members_and_classes/sealed_class_semantics_12
class FinalClass_12 {
    [int]$Val = 12
    [int]Compute() { return $this.Val * 2 }
}
$inst = [FinalClass_12]::new()
if ($inst.Compute() -ne (12 * 2)) { Write-Host "FAIL: Final class failed"; exit 1 }
Write-Host "PASS"; exit 0
