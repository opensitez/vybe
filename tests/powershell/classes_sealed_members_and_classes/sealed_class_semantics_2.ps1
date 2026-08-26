# vybe-test: powershell/classes_sealed_members_and_classes/sealed_class_semantics_2
class FinalClass_2 {
    [int]$Val = 2
    [int]Compute() { return $this.Val * 2 }
}
$inst = [FinalClass_2]::new()
if ($inst.Compute() -ne (2 * 2)) { Write-Host "FAIL: Final class failed"; exit 1 }
Write-Host "PASS"; exit 0
