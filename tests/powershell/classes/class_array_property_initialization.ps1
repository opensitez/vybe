# vybe-test: powershell/classes/class_array_property_initialization
# A PowerShell class can define strongly-typed array properties with default collection literals
class BufferPayload {
    [int[]]$Indices = @(10, 20, 30, 40)
}

$buf = [BufferPayload]::new()

if ($buf.Indices.Length -ne 4) {
    Write-Host "FAIL: expected array length 4, got $($buf.Indices.Length)"
    exit 1
}

$sum = 0
foreach ($val in $buf.Indices) {
    $sum += $val
}

if ($sum -ne 100) {
    Write-Host "FAIL: expected sum 100, got $sum"
    exit 1
}

Write-Host "PASS"
exit 0
