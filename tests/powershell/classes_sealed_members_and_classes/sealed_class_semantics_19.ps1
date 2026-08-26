# vybe-test: powershell/classes_sealed_members_and_classes/sealed_class_semantics_19
class FinalClass_19 {
    [int]$Val = 19
    [int]Compute() { return $this.Val * 2 }
}
$inst = [FinalClass_19]::new()
if ($inst.Compute() -ne (19 * 2)) { Write-Host "FAIL: Final class failed"; exit 1 }
Write-Host "PASS"; exit 0
