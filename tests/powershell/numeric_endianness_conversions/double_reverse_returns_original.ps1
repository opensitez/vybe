# vybe-test: powershell/numeric_endianness_conversions/double_reverse_returns_original
[int32]$orig = 987654321
$r1 = [System.Buffers.Binary.BinaryPrimitives]::ReverseEndianness($orig)
$r2 = [System.Buffers.Binary.BinaryPrimitives]::ReverseEndianness($r1)
if ($orig -ne $r2) {
    Write-Host "FAIL: Double reverse must return original value"
    exit 1
}
Write-Host "PASS"
exit 0
