# vybe-test: powershell/collections_generic_stack/trimexcess_on_stack
$s = [System.Collections.Generic.Stack[int]]::new(100)
$s.Push(1); $s.Push(2)
$s.TrimExcess()
if ($s.Count -ne 2) {
    Write-Host "FAIL: TrimExcess failed on stack"
    exit 1
}
Write-Host "PASS"
exit 0
