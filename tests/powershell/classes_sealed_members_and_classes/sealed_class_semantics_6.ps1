# vybe-test: powershell/classes_sealed_members_and_classes/sealed_class_semantics_6
class FinalClass_6 {
    [int]$Val = 6
    [int]Compute() { return $this.Val * 2 }
}
$inst = [FinalClass_6]::new()
if ($inst.Compute() -ne (6 * 2)) { Write-Host "FAIL: Final class failed"; exit 1 }
Write-Host "PASS"; exit 0
