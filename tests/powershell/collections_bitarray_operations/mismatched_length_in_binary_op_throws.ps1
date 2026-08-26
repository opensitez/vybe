# vybe-test: powershell/collections_bitarray_operations/mismatched_length_in_binary_op_throws
$ba1 = [System.Collections.BitArray]::new(4)
$ba2 = [System.Collections.BitArray]::new(8)
$caught = $false
try {
    $ba1.And($ba2)
} catch [System.ArgumentException] {
    $caught = $true
}
if (-not $caught) {
    Write-Host "FAIL: Expected ArgumentException on mismatched lengths"
    exit 1
}
Write-Host "PASS"
exit 0
