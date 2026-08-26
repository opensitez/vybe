# vybe-test: powershell/classes_sealed_members_and_classes/sealed_class_semantics_18
class FinalClass_18 {
    [int]$Val = 18
    [int]Compute() { return $this.Val * 2 }
}
$inst = [FinalClass_18]::new()
if ($inst.Compute() -ne (18 * 2)) { Write-Host "FAIL: Final class failed"; exit 1 }
Write-Host "PASS"; exit 0
