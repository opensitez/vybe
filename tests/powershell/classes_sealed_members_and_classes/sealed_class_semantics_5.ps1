# vybe-test: powershell/classes_sealed_members_and_classes/sealed_class_semantics_5
class FinalClass_5 {
    [int]$Val = 5
    [int]Compute() { return $this.Val * 2 }
}
$inst = [FinalClass_5]::new()
if ($inst.Compute() -ne (5 * 2)) { Write-Host "FAIL: Final class failed"; exit 1 }
Write-Host "PASS"; exit 0
