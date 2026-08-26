# vybe-test: powershell/classes_sealed_members_and_classes/sealed_class_semantics_11
class FinalClass_11 {
    [int]$Val = 11
    [int]Compute() { return $this.Val * 2 }
}
$inst = [FinalClass_11]::new()
if ($inst.Compute() -ne (11 * 2)) { Write-Host "FAIL: Final class failed"; exit 1 }
Write-Host "PASS"; exit 0
