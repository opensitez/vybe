# vybe-test: powershell/pipeline/pipeline_chain_multiple
$result = 1..20 |
    Where-Object { $_ % 2 -eq 0 } |
    ForEach-Object { $_ * $_ } |
    Measure-Object -Sum
# Even numbers 2..20: squares are 4+16+36+64+100+144+196+256+324+400 = 1540
if ($result.Sum -ne 1540) {
    Write-Host "FAIL: expected 1540, got $($result.Sum)"
    exit 1
}
Write-Host "PASS"
exit 0
