# vybe-test: powershell/collections_generic_stack/trypop_success_and_empty
$s = [System.Collections.Generic.Stack[int]]::new()
$s.Push(50)
$hasItem = ($s.Count -gt 0)
$item = if ($hasItem) { $s.Pop() } else { 0 }
if (-not $hasItem -or $item -ne 50 -or $s.Count -ne 0) {
    Write-Host "FAIL: Stack conditional pop failed"
    exit 1
}
Write-Host "PASS"
exit 0
