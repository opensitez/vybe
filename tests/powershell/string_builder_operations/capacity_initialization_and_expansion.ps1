# vybe-test: powershell/string_builder_operations/capacity_initialization_and_expansion
$sb = [System.Text.StringBuilder]::new(4)
$null = $sb.Append("12345678")
if ($sb.Capacity -lt 8 -or $sb.Length -ne 8) {
    Write-Host "FAIL: Capacity auto-expansion failed"
    exit 1
}
Write-Host "PASS"
exit 0
