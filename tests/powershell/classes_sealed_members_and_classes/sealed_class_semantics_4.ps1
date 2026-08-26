# vybe-test: powershell/classes_sealed_members_and_classes/sealed_class_semantics_4
class FinalClass_4 {
    [int]$Val = 4
    [int]Compute() { return $this.Val * 2 }
}
$inst = [FinalClass_4]::new()
if ($inst.Compute() -ne (4 * 2)) { Write-Host "FAIL: Final class failed"; exit 1 }
Write-Host "PASS"; exit 0
