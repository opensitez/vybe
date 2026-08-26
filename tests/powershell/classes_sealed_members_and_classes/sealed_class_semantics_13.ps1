# vybe-test: powershell/classes_sealed_members_and_classes/sealed_class_semantics_13
class FinalClass_13 {
    [int]$Val = 13
    [int]Compute() { return $this.Val * 2 }
}
$inst = [FinalClass_13]::new()
if ($inst.Compute() -ne (13 * 2)) { Write-Host "FAIL: Final class failed"; exit 1 }
Write-Host "PASS"; exit 0
