# vybe-test: powershell/classes_sealed_members_and_classes/sealed_class_semantics_8
class FinalClass_8 {
    [int]$Val = 8
    [int]Compute() { return $this.Val * 2 }
}
$inst = [FinalClass_8]::new()
if ($inst.Compute() -ne (8 * 2)) { Write-Host "FAIL: Final class failed"; exit 1 }
Write-Host "PASS"; exit 0
