# vybe-test: powershell/classes_sealed_members_and_classes/sealed_class_semantics_14
class FinalClass_14 {
    [int]$Val = 14
    [int]Compute() { return $this.Val * 2 }
}
$inst = [FinalClass_14]::new()
if ($inst.Compute() -ne (14 * 2)) { Write-Host "FAIL: Final class failed"; exit 1 }
Write-Host "PASS"; exit 0
